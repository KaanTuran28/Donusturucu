"""Belge dönüştürücü - profesyonel dönüşüm orkestratörü.

Desteklenen formatlar: TXT, DOCX, PDF, HTML, MD, RTF, ODT (+ Word kuruluysa
kaynak olarak eski .DOC)

Dönüşüm öncelikleri (en yüksek kaliteden düşüğe):
1. Microsoft Word (COM) kuruluysa: DOCX/DOC/RTF/ODT/HTML kaynakları Word'ün
   kendi motoruyla dönüştürülür — tablolar, şekiller, resimler, stiller
   birebir korunur.
2. PDF -> DOCX: pdf2docx ile sayfa düzeni (tablolar, resimler) korunarak.
3. Diğer tüm yollar: blok motoru (belge_bloklar) — başlıklar, kalın/italik,
   listeler ve gömülü resimler korunur.
"""

import html
import logging
import re
import shutil
import sys
import tempfile
import threading
import time
import urllib.parse
import zipfile
from contextlib import contextmanager
from pathlib import Path

from . import belge_bloklar as bloklar
from . import ortak

FORMATLAR = ["pdf", "docx", "txt", "html", "md", "rtf", "odt"]

# Word SaveAs2 format kodları (wdSaveFormat)
_WORD_FORMAT = {"docx": 16, "pdf": 17, "rtf": 6, "txt": 7, "html": 10, "odt": 23}
_WORD_ACABILIR = {"docx", "doc", "rtf", "odt", "html"}
_SADECE_WORD = {"doc"}  # dahili motorun okuyamadığı, yalnızca Word'le açılan kaynaklar
KAYNAK_FORMATLARI = set(bloklar.OKUYUCULAR) | _SADECE_WORD

# Word'ün kilitlenmesine karşı süre sınırı: taban + dosya boyutuna göre ek süre
_WORD_SURE_TABAN = 120.0     # saniye
_WORD_SURE_MB_BASINA = 30.0  # saniye / MB
_WORD_SURE_AZAMI = 1800.0

_word_durumu: bool | None = None
# DispatchEx + süreç tespiti aynı anda yalnızca bir iş parçacığında yapılır
# (iki sekme aynı anda Word başlatırsa bekçiler birbirinin Word'ünü sanmasın)
_WORD_BASLATMA_KILIDI = threading.Lock()


def _word_kullanilabilir() -> bool:
    """Microsoft Word'ün COM üzerinden kullanılabilir olup olmadığını denetler."""
    global _word_durumu
    if _word_durumu is None:
        try:
            import winreg

            import win32com.client  # noqa: F401

            winreg.CloseKey(winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT, "Word.Application"))
            _word_durumu = True
        except Exception:
            _word_durumu = False
    return _word_durumu


# ============================================================ Word süreç bekçisi

def _word_pidleri() -> set[int]:
    """Çalışan WINWORD.EXE süreçlerinin kimlikleri (Toolhelp anlık görüntüsü)."""
    if sys.platform != "win32":
        return set()
    import ctypes
    from ctypes import wintypes

    class SurecGirdisi(ctypes.Structure):
        _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD),
                    ("th32ProcessID", wintypes.DWORD),
                    ("th32DefaultHeapID", ctypes.c_size_t),
                    ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
                    ("th32ParentProcessID", wintypes.DWORD),
                    ("pcPriClassBase", ctypes.c_long), ("dwFlags", wintypes.DWORD),
                    ("szExeFile", ctypes.c_wchar * 260)]

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    k32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
    k32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(SurecGirdisi)]
    k32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(SurecGirdisi)]
    k32.CloseHandle.argtypes = [wintypes.HANDLE]

    anlik = k32.CreateToolhelp32Snapshot(0x2, 0)  # TH32CS_SNAPPROCESS
    if not anlik or anlik == ctypes.c_void_p(-1).value:
        return set()
    pidler = set()
    try:
        girdi = SurecGirdisi()
        girdi.dwSize = ctypes.sizeof(SurecGirdisi)
        var = k32.Process32FirstW(anlik, ctypes.byref(girdi))
        while var:
            if girdi.szExeFile.lower() == "winword.exe":
                pidler.add(girdi.th32ProcessID)
            var = k32.Process32NextW(anlik, ctypes.byref(girdi))
    finally:
        k32.CloseHandle(anlik)
    return pidler


def _sureci_sonlandir(pid: int):
    import ctypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.OpenProcess.restype = ctypes.c_void_p
    tutamac = k32.OpenProcess(0x0001, False, pid)  # PROCESS_TERMINATE
    if tutamac:
        k32.TerminateProcess(ctypes.c_void_p(tutamac), 1)
        k32.CloseHandle(ctypes.c_void_p(tutamac))


# Word'ün iletişim kutusu pencere sınıfları
_DIYALOG_SINIFLARI = {"NUIDialog", "#32770", "bosa_sdm_msword"}
_AZAMI_DIYALOG = 5


def _diyaloglari_kapat(pidler: set[int]) -> int:
    """Bizim Word sürecimizin açtığı görünür iletişim kutularını kapatır
    (Kapat = İptal). Görünmez çalışan Word'de hiçbir iletişim kutusu
    beklenmez; açılan her kutu (parola sorma, onarım, dosya kullanımda...)
    Word'ü sonsuza kadar bekletir ve kullanıcının ekranında belirir."""
    import win32con
    import win32gui
    import win32process

    kapatilan = []

    def bak(pencere, _):
        try:
            _, pid = win32process.GetWindowThreadProcessId(pencere)
            if (pid in pidler and win32gui.IsWindowVisible(pencere)
                    and win32gui.GetClassName(pencere) in _DIYALOG_SINIFLARI):
                win32gui.PostMessage(pencere, win32con.WM_CLOSE, 0, 0)
                kapatilan.append(pencere)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(bak, None)
    except Exception:
        pass
    return len(kapatilan)


def _bekci_baslat(pidler: set[int], sure: float, iptal) -> tuple[threading.Event, dict]:
    """Bizim başlattığımız Word sürecini izleyen bekçi:
    - beklenmedik iletişim kutularını kapatır (bekleyen çağrı hata ile döner),
    - çok sayıda kutu açılırsa, süre dolarsa ya da kullanıcı iptal ederse
      süreci sonlandırır.
    Böylece Word hiçbir durumda programı sonsuza kadar kilitleyemez."""
    bitti = threading.Event()
    durum = {"neden": None}

    def izle():
        son = time.monotonic() + sure
        diyalog_sayisi = 0
        while not bitti.wait(0.5):
            diyalog_sayisi += _diyaloglari_kapat(pidler)
            if iptal is not None and iptal.is_set():
                durum["neden"] = "iptal"
            elif time.monotonic() > son or diyalog_sayisi > _AZAMI_DIYALOG:
                durum["neden"] = "zaman"
            else:
                continue
            for pid in pidler:
                _sureci_sonlandir(pid)
            return

    threading.Thread(target=izle, daemon=True).start()
    return bitti, durum


_OLE_IMZASI = bloklar._OLE_IMZASI


def parolali_mi(yol: Path) -> bool:
    """Belgenin açılış parolasıyla şifreli olup olmadığını, Word'e vermeden
    tespit eder (Word'e verilince ekranda parola kutusu açıp takılıyordu).

    - Şifreli DOCX: ZIP değil OLE kabıdır ve "EncryptionInfo" akışı içerir.
    - Eski .DOC: WordDocument akışındaki FIB başlığında fEncrypted biti.
    - ODT: manifest.xml'de "encryption-data" kaydı.
    """
    try:
        with open(yol, "rb") as f:
            imza = f.read(8)
    except OSError:
        return False
    if imza == _OLE_IMZASI:
        try:
            import pythoncom

            okuma = 0x00000000 | 0x00000020  # STGM_READ | STGM_SHARE_DENY_WRITE
            depo = pythoncom.StgOpenStorageEx(
                str(yol), okuma, 0, 0, pythoncom.IID_IStorage)
            adlar = {oge[0] for oge in depo.EnumElements()}
            if {"EncryptionInfo", "EncryptedPackage"} & adlar:
                return True
            if "WordDocument" in adlar:
                akis = depo.OpenStream("WordDocument", None, 0x10, 0)  # SHARE_EXCLUSIVE
                fib = akis.Read(12)
                return bool(int.from_bytes(fib[10:12], "little") & 0x0100)
        except Exception:
            return False
    elif imza[:4] == b"PK\x03\x04" and yol.suffix.lower() == ".odt":
        try:
            with zipfile.ZipFile(yol) as arsiv:
                return b"encryption-data" in arsiv.read("META-INF/manifest.xml")
        except Exception:
            return False
    return False


def _resimleri_gom(belge):
    """HTML'den açılan belgelerde Word resimleri dış bağlantı olarak tutar;
    bağlantıyı koparıp resmi belgeye gömer. (Eskiden HTML -> DOCX çıktısındaki
    resimler yalnızca diskteki dosyaya bağlantıydı; belge başka bilgisayara
    gönderilince resimler kayboluyordu.)"""
    for koleksiyon in (belge.InlineShapes, belge.Shapes):
        try:
            adet = koleksiyon.Count
        except Exception:
            continue
        for sira in range(adet, 0, -1):
            try:
                baglanti = koleksiyon.Item(sira).LinkFormat
                if baglanti is not None:
                    _baglanti_yolunu_onar(baglanti)
                    baglanti.SavePictureWithDocument = True
                    baglanti.BreakLink()
            except Exception:
                pass  # bağlantısız şekil / desteklenmeyen tür


def _baglanti_yolunu_onar(baglanti):
    """Word yüzde-kodlu yolları ("resimler%20klas%C3%B6r%C3%BC/a.png")
    çözemiyor; çözülmüş yol diskte varsa bağlantıyı ona çevirip yeniler."""
    try:
        kaynak = baglanti.SourceFullName or ""
        if "%" not in kaynak:
            return
        cozulmus = urllib.parse.unquote(kaynak)
        if cozulmus.replace("/", "\\").startswith("\\\\") or not Path(cozulmus).is_file():
            return
        baglanti.SourceFullName = cozulmus
        baglanti.Update()
    except Exception:
        pass


@contextmanager
def _word_icin_hazirla(kaynak: Path):
    """Word'e verilecek dosyayı hazırlar.

    Karakter kümesi bildirmeyen UTF-8 HTML'i Word ANSI sanıp Türkçe metni ve
    Türkçe karakterli resim yollarını bozuyordu (Documents.Open'ın Encoding
    ayarını HTML'de yok sayıyor). Böyle dosyaların geçici bir kopyasına
    <meta charset="utf-8"> ve göreli adresler yine kaynak klasörden çözülsün
    diye <base href> eklenir. Diğer tüm dosyalar olduğu gibi verilir."""
    veri = kaynak.read_bytes() if kaynak.suffix.lower() in (".html", ".htm") else b""
    bildirimsiz_utf8 = False
    if veri and not veri.startswith((b"\xef\xbb\xbf", b"\xff\xfe", b"\xfe\xff")) \
            and not re.search(rb"<meta[^>]+charset", veri[:4096], re.I):
        try:
            veri.decode("utf-8")
            bildirimsiz_utf8 = True
        except UnicodeDecodeError:
            pass
    if not bildirimsiz_utf8:
        yield kaynak
        return

    metin = veri.decode("utf-8")
    ek = '<meta charset="utf-8">'
    if not re.search(r"<base\b", metin, re.I):
        ek += f'<base href="{html.escape(str(kaynak.parent.resolve()))}\\">'
    eslesme = re.search(r"<head\b[^>]*>", metin, re.I) or re.search(r"<html\b[^>]*>", metin, re.I)
    konum = eslesme.end() if eslesme else 0
    with tempfile.TemporaryDirectory() as gecici:
        kopya = Path(gecici) / kaynak.name
        kopya.write_text(metin[:konum] + ek + metin[konum:], encoding="utf-8")
        yield kopya


def _word_ile_donustur(kaynak: Path, cikti: Path, hedef_format: str, iptal=None):
    """Belgeyi Microsoft Word'ün kendi motoruyla dönüştürür (tam kalite)."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    belge = None
    bitti = None
    durum = {"neden": None}
    try:
        with _WORD_BASLATMA_KILIDI:
            onceki = _word_pidleri()
            word = win32com.client.DispatchEx("Word.Application")
            bizim = _word_pidleri() - onceki
        sure = min(_WORD_SURE_AZAMI, _WORD_SURE_TABAN
                   + kaynak.stat().st_size / 1e6 * _WORD_SURE_MB_BASINA)
        bitti, durum = _bekci_baslat(bizim, sure, iptal)

        word.Visible = False
        word.DisplayAlerts = 0
        # Makrolar hiçbir koşulda çalışmasın (msoAutomationSecurityForceDisable).
        # Otomasyonla açılan belgelerde varsayılan "makrolara izin ver"dir.
        word.AutomationSecurity = 3
        belge = word.Documents.Open(
            str(kaynak.resolve()),
            ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
            # Sahte parola: bazı durumlarda parola kutusu yerine anında hata
            # döndürür (parolasız belgelerde yok sayılır). Asıl koruma
            # parolali_mi() ön denetimi ve bekçinin kutu kapatmasıdır.
            PasswordDocument="\u0001yanlis-parola",
            Visible=False, NoEncodingDialog=True)
        if kaynak.suffix.lower() in (".html", ".htm"):
            _resimleri_gom(belge)
        belge.SaveAs2(str(cikti.resolve()), FileFormat=_WORD_FORMAT[hedef_format])
    except Exception as hata:
        if durum["neden"] == "iptal":
            raise ortak.IptalEdildi() from None
        if durum["neden"] == "zaman":
            raise RuntimeError("Microsoft Word yanıt vermedi (zaman aşımı).") from hata
        raise
    finally:
        # Word zaten kapanmış/sonlandırılmış olabilir: kapatma hataları yok
        # sayılır. Bekçi kapatma bitene kadar açık kalır (Quit de takılabilir).
        if belge is not None:
            try:
                belge.Close(0)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(0)  # wdDoNotSaveChanges
            except Exception:
                pass
        if bitti is not None:
            bitti.set()
        pythoncom.CoUninitialize()

    if hedef_format == "txt":
        _utf8e_cevir(cikti)


def _utf8e_cevir(yol: Path):
    """Word'ün kaydettiği metni (ANSI/UTF-16 olabilir) UTF-8'e normalize eder."""
    veri = yol.read_bytes()
    if veri.startswith((b"\xff\xfe", b"\xfe\xff")):
        metin = veri.decode("utf-16")
    else:
        try:
            metin = veri.decode("utf-8-sig")
        except UnicodeDecodeError:
            # "mbcs" = Windows'un gerçek ANSI kod sayfası (UTF-8 modundan etkilenmez)
            metin = veri.decode("mbcs", errors="replace")
    yol.write_text(metin, encoding="utf-8")


def _pdf2docx_ile(kaynak: Path, cikti: Path):
    """PDF'i sayfa düzenini (tablo/resim/sütun) koruyarak DOCX'e çevirir."""
    from pdf2docx import Converter

    logging.disable(logging.INFO)
    try:
        # password="": kullanıcı parolası olmayan (yalnızca sahip/izin
        # parolalı) PDF'lerin çoğu bununla otomatik açılır. Gerçek bir
        # kullanıcı parolası gerekiyorsa yine de hata verir; donustur() bu
        # durumda blok motoruna düşer.
        donusturucu = Converter(str(kaynak), password="")
        try:
            donusturucu.convert(str(cikti))
        finally:
            donusturucu.close()
    finally:
        logging.disable(logging.NOTSET)


def _md_html_yaz(kaynak_yol: Path, gecici: Path) -> Path:
    """Markdown'ı Word'ün açacağı geçici bir HTML dosyasına çevirir."""
    govde = bloklar.md_html(bloklar.metin_coz(kaynak_yol.read_bytes()))
    govde = bloklar.resim_adreslerini_yerellestir(govde, kaynak_yol.parent, gecici)
    ara_html = gecici / (kaynak_yol.stem + ".html")
    ara_html.write_text(
        '<html><head><meta charset="utf-8"></head><body>'
        + govde + "</body></html>", encoding="utf-8")
    return ara_html


def _blok_motoru(kaynak: Path, kaynak_format: str, cikti: Path, hedef_format: str):
    liste = bloklar.OKUYUCULAR[kaynak_format](kaynak)
    bloklar.YAZICILAR[hedef_format](liste, cikti)


def donustur(kaynak: str, hedef_format: str, cikti_klasoru: str | None = None,
             iptal=None) -> str:
    """Bir belgeyi hedef formata dönüştürür, çıktı yolunu döndürür."""
    kaynak_yol = Path(kaynak)
    kaynak_format = kaynak_yol.suffix.lower().lstrip(".")
    if kaynak_format == "htm":
        kaynak_format = "html"
    hedef_format = hedef_format.lower().lstrip(".")

    if kaynak_format not in KAYNAK_FORMATLARI:
        raise ValueError(f"Desteklenmeyen kaynak belge formatı: {kaynak_format}")
    if hedef_format not in bloklar.YAZICILAR:
        raise ValueError(f"Desteklenmeyen hedef belge formatı: {hedef_format}")
    if kaynak_format in _SADECE_WORD and not _word_kullanilabilir():
        raise RuntimeError(
            f".{kaynak_format} dosyalarını dönüştürmek için bilgisayarda "
            "Microsoft Word kurulu olmalı.")
    ortak.kaynak_denetle(kaynak_yol)
    if kaynak_format in ("docx", "doc", "odt") and parolali_mi(kaynak_yol):
        raise RuntimeError(
            "Bu belge parola ile korunuyor; parola olmadan dönüştürülemiyor.")

    with ortak.cikti_dosyasi(kaynak_yol, hedef_format, cikti_klasoru) as cikti:
        if kaynak_format == hedef_format:
            # Aynı format: hiçbir motordan geçirmeden kayıpsız kopyala.
            # (Word/blok motoru ile "yeniden kaydetmek" PDF gibi formatlarda
            # içeriği düz metne indirger; doğrudan kopya hem daha hızlı hem
            # her zaman kayıpsız.)
            shutil.copy2(kaynak_yol, cikti)
        else:
            _donustur_yola(kaynak_yol, kaynak_format, cikti, hedef_format, iptal)
    return str(cikti)


def _donustur_yola(kaynak_yol: Path, kaynak_format: str, cikti: Path,
                   hedef_format: str, iptal=None):
    """Kaynağı, önceden ayrılmış `cikti` yoluna en kaliteli yolla yazar."""
    # ---- 1) PDF kaynaklı dönüşümler: pdf2docx ile düzen korunur
    if kaynak_format == "pdf" and hedef_format == "docx":
        try:
            _pdf2docx_ile(kaynak_yol, cikti)
            return
        except Exception:
            pass  # şifreli/bozuk/karmaşık PDF -- blok motoruna düş (aşağıda)
    if kaynak_format == "pdf" and hedef_format in ("html", "rtf", "odt"):
        # PDF -> (düzen korumalı) DOCX -> hedef
        with tempfile.TemporaryDirectory() as gecici:
            ara_docx = Path(gecici) / (kaynak_yol.stem + ".docx")
            try:
                _pdf2docx_ile(kaynak_yol, ara_docx)
                _donustur_yola(ara_docx, "docx", cikti, hedef_format, iptal)
                return
            except ortak.IptalEdildi:
                raise
            except Exception:
                pass  # blok motoruna düş

    # ---- 2) Microsoft Word yolu (tam kalite)
    # Ağ (UNC) adresine başvuran HTML Word'e verilmez (bkz. ag_adresi_var).
    word_uygun = _word_kullanilabilir() and not (
        kaynak_format == "html" and bloklar.ag_adresi_var(
            bloklar.metin_coz(kaynak_yol.read_bytes(), html_mi=True)))
    if word_uygun:
        try:
            if kaynak_format in _WORD_ACABILIR and hedef_format in _WORD_FORMAT:
                with _word_icin_hazirla(kaynak_yol) as word_kaynagi:
                    _word_ile_donustur(word_kaynagi, cikti, hedef_format, iptal)
                return
            if kaynak_format in _WORD_ACABILIR and kaynak_format not in ("docx", "html"):
                # Word'ün yazamadığı hedef (MD): Word ile DOCX'e, oradan blok
                # motoruna -- RTF/ODT/DOC'taki başlık/kalın/resimler korunur.
                with tempfile.TemporaryDirectory() as gecici:
                    ara_docx = Path(gecici) / (kaynak_yol.stem + ".docx")
                    _word_ile_donustur(kaynak_yol, ara_docx, "docx", iptal)
                    _blok_motoru(ara_docx, "docx", cikti, hedef_format)
                return
            if kaynak_format == "md" and hedef_format in _WORD_FORMAT and hedef_format != "html":
                # MD -> HTML (markdown) -> Word ile hedefe
                with tempfile.TemporaryDirectory() as gecici:
                    ara_html = _md_html_yaz(kaynak_yol, Path(gecici))
                    if not bloklar.ag_adresi_var(ara_html.read_text(encoding="utf-8")):
                        _word_ile_donustur(ara_html, cikti, hedef_format, iptal)
                        return
                # MD içindeki ham HTML ağ adresine başvuruyor: dahili motora düş
        except ortak.IptalEdildi:
            raise
        except Exception as hata:
            if kaynak_format in _SADECE_WORD:
                raise RuntimeError(ortak.hata_metni(hata)) from hata
            # Word başarısız olursa blok motoruna düş

    # ---- 3) MD -> HTML: gerçek Markdown işleme
    if kaynak_format == "md" and hedef_format == "html":
        liste = bloklar.oku_md(kaynak_yol)
        bloklar.yaz_html(liste, cikti)
        return

    # ---- 4) Blok motoru (başlık/kalın/italik/liste/resim korumalı)
    _blok_motoru(kaynak_yol, kaynak_format, cikti, hedef_format)
