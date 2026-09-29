"""Tüm dönüştürücülerin paylaştığı yardımcılar.

- Çıktı dosyası adı seçimi: hiçbir zaman var olan bir dosyanın (kaynağın
  kendisi ya da aynı toplu işteki başka bir kaynak dahil) üzerine yazmaz.
- Dönüştürme başarısız olursa / iptal edilirse yarım kalan çıktıyı siler.
- ffmpeg'i güvenli ayarlarla, iptal edilebilir biçimde çalıştırır.
- Kütüphanelerin İngilizce/teknik hatalarını anlaşılır Türkçe mesaja çevirir.
"""

import errno
import re
import subprocess
import sys
from contextlib import contextmanager
from pathlib import Path

# Windows'ta konsol penceresi açılmasını engelle
_CREATE_NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0


class IptalEdildi(Exception):
    """Kullanıcı dönüştürmeyi iptal etti."""

    def __init__(self):
        super().__init__("Dönüştürme iptal edildi.")


def kaynak_denetle(kaynak_yol: Path):
    if not kaynak_yol.is_file():
        raise FileNotFoundError(errno.ENOENT, "Dosya bulunamadı", str(kaynak_yol))


def _yer_ayir(kaynak_yol: Path, hedef_format: str, klasor: Path) -> Path:
    """Boş bir çıktı adı seçer ve dosyayı oluşturarak o adı ayırır.

    Sıra: ad.uzanti -> ad_donusturuldu.uzanti -> ad_donusturuldu_2.uzanti ...
    Var olan hiçbir dosya ezilmez. (Eskiden toplu işte "a.png" ile "a.jpg"
    birlikte JPG'ye çevrilince, a.png'nin çıktısı henüz işlenmemiş kaynak
    a.jpg'nin üzerine yazıyordu.)
    """
    for sira in range(1, 10_000):
        if sira == 1:
            ek = ""
        elif sira == 2:
            ek = "_donusturuldu"
        else:
            ek = f"_donusturuldu_{sira - 1}"
        aday = klasor / f"{kaynak_yol.stem}{ek}.{hedef_format}"
        try:
            with open(aday, "xb"):  # "x": dosya varsa hata -> yarış durumunda da güvenli
                pass
            return aday
        except FileExistsError:
            continue
    raise RuntimeError("Çıktı için boş bir dosya adı bulunamadı.")


@contextmanager
def cikti_dosyasi(kaynak_yol: Path, hedef_format: str, cikti_klasoru: str | None):
    """Çıktı yolunu ayırır; blok hata/iptal ile biterse yarım dosyayı siler."""
    klasor = Path(cikti_klasoru) if cikti_klasoru else kaynak_yol.parent
    klasor.mkdir(parents=True, exist_ok=True)
    cikti = _yer_ayir(kaynak_yol, hedef_format, klasor)
    try:
        yield cikti
    except BaseException:
        try:
            cikti.unlink(missing_ok=True)
        except OSError:
            pass
        raise


# ================================================================== ffmpeg

def ffmpeg_hata_metni(stderr: str, akis: str = "ses") -> str:
    """ffmpeg stderr çıktısından anlamlı/açıklayıcı bir hata satırı seçer.

    Son satırı almak genelde "Conversion failed!" ya da ilerleme istatistiği
    ("frame= 0 fps=0.0 ...") gibi anlamsız bir satır verir; gerçek nedeni
    taşıyan satır genelde ondan öncedir.
    """
    metin = stderr or ""
    if "does not contain any stream" in metin or "matches no streams" in metin:
        return f"Kaynak dosyada beklenen {akis} akışı bulunamadı."
    if "Invalid data found when processing input" in metin:
        return "Kaynak dosya okunamadı (bozuk veya desteklenmeyen bir biçim)."
    if "No such file or directory" in metin:
        return "Kaynak dosya bulunamadı."
    if "No space left on device" in metin:
        return "Diskte yer kalmadı."
    satirlar = [s.strip() for s in metin.strip().splitlines() if s.strip()]
    for satir in reversed(satirlar):
        # "[wmav2 @ 000001b3...] too many channels" -> "too many channels"
        sade = re.sub(r"^\[[^\]]*@ [0-9a-fA-Fx]+\]\s*", "", satir)
        if not sade or any(k in sade for k in _FFMPEG_GENEL_SATIRLAR):
            continue
        return sade
    return satirlar[-1] if satirlar else "bilinmeyen hata"


# ffmpeg 7'nin her hatada bastığı, asıl nedeni söylemeyen genel satırlar
_FFMPEG_GENEL_SATIRLAR = (
    "Conversion failed!", "Error opening output file", "Task finished with error code",
    "Terminating thread with return code", "Nothing was written into output file",
    "Could not open encoder before EOF", "Error sending frames to consumers",
    "Error while opening encoder", "Error while filtering", "Exiting normally",
    "frame=", "size=", "video:", "Press [q] to stop",
)


def ffmpeg_calistir(kaynak_yol: Path, ayarlar: list[str], cikti: Path,
                    iptal=None, akis: str = "ses"):
    """ffmpeg ile kaynak -> çıktı dönüşümü yapar; iptal edilirse süreci öldürür.

    Güvenlik: "-protocol_whitelist file" ile ffmpeg yalnızca yerel dosya
    okuyabilir; kötü niyetli bir medya dosyası (ör. içine gömülü bir HLS
    oynatma listesi) ffmpeg'e internetten / ağ paylaşımından veri çektiremez.
    "file:" öneki, dosya adının bir protokol adı gibi yorumlanmasını engeller.
    """
    import imageio_ffmpeg

    komut = [
        imageio_ffmpeg.get_ffmpeg_exe(),
        "-hide_banner", "-nostdin", "-nostats", "-y",
        "-protocol_whitelist", "file",
        "-i", "file:" + str(kaynak_yol),
        *ayarlar,
        "file:" + str(cikti),
    ]
    with subprocess.Popen(
        komut,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=_CREATE_NO_WINDOW,
    ) as surec:
        while True:
            try:
                _, stderr = surec.communicate(timeout=0.25)
                break
            except subprocess.TimeoutExpired:
                if iptal is not None and iptal.is_set():
                    surec.kill()
                    surec.communicate()
                    raise IptalEdildi() from None
    if surec.returncode != 0:
        raise RuntimeError(f"ffmpeg hatası: {ffmpeg_hata_metni(stderr, akis)}")


# ============================================================ hata mesajları

def hata_metni(hata: BaseException) -> str:
    """Bir istisnayı kullanıcıya gösterilecek anlaşılır Türkçe metne çevirir."""
    ad = type(hata).__name__
    if isinstance(hata, FileNotFoundError):
        return "Dosya bulunamadı (taşınmış ya da silinmiş olabilir)."
    if isinstance(hata, PermissionError):
        return ("Dosyaya erişilemedi: başka bir programda açık olabilir ya da "
                "bu klasöre yazma izniniz yok.")
    if isinstance(hata, OSError) and hata.errno == errno.ENOSPC:
        return "Diskte yer kalmadı."
    if isinstance(hata, MemoryError):
        return "Bellek yetersiz; dosya bu bilgisayarda işlenemeyecek kadar büyük."
    if ad == "UnidentifiedImageError":
        return "Görsel tanınamadı (dosya bozuk ya da desteklenmeyen bir biçimde)."
    if ad == "DecompressionBombError":
        return "Görsel çok büyük (güvenlik sınırı aşıldı); işlenmedi."
    if ad in ("BadZipFile", "PackageNotFoundError"):
        return "Dosya bozuk ya da uzantısıyla uyuşmayan bir biçimde."
    if ad in ("XMLSyntaxError", "ExpatError", "SAXParseException"):
        return "Belge bozuk; içeriği okunamadı."
    if ad in ("PdfReadError", "PdfStreamError", "EmptyFileError"):
        return "PDF okunamadı (dosya bozuk olabilir)."
    if ad == "com_error":
        try:
            aciklama = hata.args[2][2]
            if aciklama:
                return "Microsoft Word: " + " ".join(str(aciklama).split())
        except (IndexError, TypeError):
            pass
        return "Microsoft Word işlemi başarısız oldu."
    metin = str(hata).strip()
    return metin or ad
