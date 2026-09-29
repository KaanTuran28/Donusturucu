"""Dönüştürücü - Belge / Görsel / Ses / Video dosya dönüştürme programı.

Çalıştırma: python donusturucu.py  (veya Başlat.bat'a çift tıklayın)
"""

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import sv_ttk

from converters import belge, gorsel, ses, video
from converters.ortak import IptalEdildi, hata_metni

try:  # sürükle-bırak isteğe bağlıdır: paket yoksa program yine çalışır
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = TkinterDnD = None


def _desen(uzantilar) -> str:
    return " ".join(f"*.{u}" for u in uzantilar)


_BELGE = ("pdf", "docx", "doc", "txt", "html", "htm", "md", "rtf", "odt")
_GORSEL = ("png", "jpg", "jpeg", "bmp", "gif", "webp", "tiff", "tif", "ico",
           "avif", "heic", "heif")
_SES = ("mp3", "wav", "ogg", "flac", "m4a", "aac", "wma", "opus", "aiff", "aif")
_VIDEO = ("mp4", "avi", "mkv", "mov", "webm", "wmv", "flv", "m4v", "mpg", "mpeg",
          "3gp", "ts")

SEKMELER = [
    {
        "ad": "📄 Belgeler",
        "modul": belge,
        "formatlar": belge.FORMATLAR,
        "kabul": set(_BELGE),
        "uzantilar": [("Belgeler", _desen(_BELGE))],
    },
    {
        "ad": "🖼️ Görseller",
        "modul": gorsel,
        "formatlar": gorsel.FORMATLAR,
        "kabul": set(_GORSEL),
        "uzantilar": [("Görseller", _desen(_GORSEL))],
    },
    {
        "ad": "🎵 Sesler",
        "modul": ses,
        "formatlar": ses.FORMATLAR,
        "kabul": set(_SES + _VIDEO),
        "uzantilar": [("Ses dosyaları", _desen(_SES)),
                      ("Video (ses çıkarma)", _desen(_VIDEO))],
    },
    {
        "ad": "🎬 Videolar",
        "modul": video,
        "formatlar": video.FORMATLAR,
        "kabul": set(_VIDEO + ("gif",)),
        "uzantilar": [("Video dosyaları", _desen(_VIDEO + ("gif",)))],
    },
]

_AZAMI_KLASOR_DOSYASI = 5000  # klasör eklemede güvenlik sınırı
_VARSAYILAN_CIKTI = "(kaynak dosyanın yanına)"

# ============================================================== tema / kaynak

# sv_ttk paletiyle birebir eşleşen renkler -- klasik tk.Listbox ttk temasından
# otomatik renk almadığı için elle senkronize edilir.
_LISTE_RENKLERI = {
    "light": {"arka": "#fafafa", "yazi": "#1c1c1c",
              "secili_arka": "#e7e7e7", "secili_yazi": "#191919"},
    "dark": {"arka": "#1c1c1c", "yazi": "#fafafa",
             "secili_arka": "#292929", "secili_yazi": "#fafafa"},
}

_AYAR_DOSYASI = (
    Path(os.getenv("APPDATA", str(Path.home()))) / "Dönüştürücü" / "ayarlar.json")


def _kaynak_yolu(dosya_adi: str) -> Path:
    """Kaynaktan çalışırken proje klasörünü, paketlenmiş .exe'de gömülü
    kaynak klasörünü döndürür (icon.ico gibi dosyalar için)."""
    taban = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return taban / dosya_adi


def _tema_yukle() -> str:
    try:
        veri = json.loads(_AYAR_DOSYASI.read_text(encoding="utf-8"))
        if veri.get("tema") in _LISTE_RENKLERI:
            return veri["tema"]
    except Exception:
        pass
    return "light"


def _tema_kaydet(tema: str):
    try:
        _AYAR_DOSYASI.parent.mkdir(parents=True, exist_ok=True)
        _AYAR_DOSYASI.write_text(json.dumps({"tema": tema}), encoding="utf-8")
    except Exception:
        pass  # tercih kaydedilemezse sessizce yok say, uygulamayı engellemesin


def _tema_dugme_metni(tema: str) -> str:
    return "🌙  Koyu Tema" if tema == "light" else "☀️  Açık Tema"


def _gezginde_goster(yol: str):
    """Dosyayı Windows Gezgini'nde seçili olarak gösterir."""
    try:
        if sys.platform == "win32":
            # Windows yollarında çift tırnak bulunamaz; komut enjeksiyonu yok.
            subprocess.Popen(f'explorer /select,"{os.path.normpath(yol)}"')
        else:
            subprocess.Popen(["xdg-open", str(Path(yol).parent)])
    except Exception:
        pass


class DonusturmeSekmesi(ttk.Frame):
    """Tek bir dosya türü (belge/görsel/ses/video) için dönüştürme paneli."""

    def __init__(self, ana, bilgi, surukle_birak: bool = False):
        super().__init__(ana, padding=18)
        self.bilgi = bilgi
        self.kuyruk: queue.Queue = queue.Queue()
        self.iptal = threading.Event()
        self.calisiyor = False
        self.son_cikti: str | None = None
        self._arayuzu_kur(surukle_birak)
        if surukle_birak:
            self._surukle_birak_etkinlestir()

    def _arayuzu_kur(self, surukle_birak: bool):
        # --- dosya listesi
        baslik = ttk.Frame(self)
        baslik.grid(row=0, column=0, columnspan=3, sticky="ew")
        ttk.Label(baslik, text="Dönüştürülecek dosyalar",
                  font="SunValleyBodyStrongFont").pack(side=tk.LEFT)
        if surukle_birak:
            ttk.Label(baslik, text="  —  dosya ya da klasörleri buraya sürükleyip "
                                   "bırakabilirsiniz",
                      font="SunValleyCaptionFont").pack(side=tk.LEFT)

        liste_cerceve = ttk.Frame(self)
        liste_cerceve.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(8, 10))
        self.liste = tk.Listbox(
            liste_cerceve, height=9, selectmode=tk.EXTENDED,
            relief="flat", borderwidth=0, highlightthickness=1,
            activestyle="none", font="SunValleyBodyFont")
        kaydirma = ttk.Scrollbar(liste_cerceve, command=self.liste.yview)
        self.liste.configure(yscrollcommand=kaydirma.set)
        self.liste.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        kaydirma.pack(side=tk.RIGHT, fill=tk.Y)
        self.liste.bind("<Delete>", lambda _o: self._secileni_kaldir())
        self.liste_temasini_uygula()

        dugmeler = ttk.Frame(self)
        dugmeler.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 16))
        self.liste_dugmeleri = [
            ttk.Button(dugmeler, text="➕ Dosya Ekle", command=self._dosya_ekle),
            ttk.Button(dugmeler, text="📁 Klasör Ekle", command=self._klasor_ekle),
            ttk.Button(dugmeler, text="➖ Seçileni Kaldır", command=self._secileni_kaldir),
            ttk.Button(dugmeler, text="🗑️ Temizle",
                       command=lambda: self.liste.delete(0, tk.END)),
        ]
        for dugme in self.liste_dugmeleri:
            dugme.pack(side=tk.LEFT, padx=(0, 6))

        # --- hedef format
        ttk.Label(self, text="Hedef format",
                  font="SunValleyBodyStrongFont").grid(row=3, column=0, sticky="w")
        self.format_kutusu = ttk.Combobox(
            self, values=self.bilgi["formatlar"], state="readonly", width=10)
        self.format_kutusu.current(0)
        self.format_kutusu.grid(row=4, column=0, sticky="w", padx=(0, 6), pady=(6, 16))

        # --- çıktı klasörü
        ttk.Label(self, text="Çıktı klasörü",
                  font="SunValleyBodyStrongFont").grid(
            row=3, column=1, columnspan=2, sticky="w")
        cikti_satiri = ttk.Frame(self)
        cikti_satiri.grid(row=4, column=1, columnspan=2, sticky="ew", pady=(6, 16))
        cikti_satiri.columnconfigure(0, weight=1)
        self.cikti_yolu = tk.StringVar(value=_VARSAYILAN_CIKTI)
        ttk.Entry(cikti_satiri, textvariable=self.cikti_yolu, state="readonly").grid(
            row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(cikti_satiri, text="📁 Seç...", command=self._klasor_sec).grid(
            row=0, column=1)
        ttk.Button(cikti_satiri, text="↺", width=3,
                   command=lambda: self.cikti_yolu.set(_VARSAYILAN_CIKTI)).grid(
            row=0, column=2, padx=(6, 0))

        # --- dönüştür / iptal + ilerleme
        eylem = ttk.Frame(self)
        eylem.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        eylem.columnconfigure(0, weight=1)
        self.donustur_dugmesi = ttk.Button(
            eylem, text="🔄  DÖNÜŞTÜR", style="Accent.TButton",
            command=self._donusturmeyi_baslat)
        self.donustur_dugmesi.grid(row=0, column=0, sticky="ew", ipady=6)
        self.iptal_dugmesi = ttk.Button(
            eylem, text="⏹  İptal", state="disabled", command=self._iptal_et)
        self.iptal_dugmesi.grid(row=0, column=1, padx=(8, 0), ipady=6)

        self.ilerleme = ttk.Progressbar(self, mode="determinate")
        self.ilerleme.grid(row=6, column=0, columnspan=3, sticky="ew")

        alt = ttk.Frame(self)
        alt.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        alt.columnconfigure(0, weight=1)
        self.durum = tk.StringVar(value="Hazır.")
        self.durum_etiketi = ttk.Label(alt, textvariable=self.durum, wraplength=460)
        self.durum_etiketi.grid(row=0, column=0, sticky="w")
        self.ac_dugmesi = ttk.Button(
            alt, text="📂 Çıktıyı Göster",
            command=lambda: self.son_cikti and _gezginde_goster(self.son_cikti))
        # sonuç oluşunca görünür olur
        alt.bind("<Configure>", lambda o: self.durum_etiketi.configure(
            wraplength=max(200, o.width - 170)))

        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=1)
        self.rowconfigure(1, weight=1)

    def liste_temasini_uygula(self, tema: str | None = None):
        """Klasik tk.Listbox'ın renklerini geçerli sv_ttk temasıyla eşitler."""
        renk = _LISTE_RENKLERI[tema or sv_ttk.get_theme()]
        self.liste.configure(
            background=renk["arka"], foreground=renk["yazi"],
            selectbackground=renk["secili_arka"], selectforeground=renk["secili_yazi"],
            highlightbackground=renk["secili_arka"], highlightcolor=renk["secili_arka"])

    # ------------------------------------------------ dosya ekleme

    def _surukle_birak_etkinlestir(self):
        for hedef in (self, self.liste):
            hedef.drop_target_register(DND_FILES)
            hedef.dnd_bind("<<Drop>>", self._birakildi)

    def _birakildi(self, olay):
        if not self.calisiyor:
            self.yollari_ekle(self.tk.splitlist(olay.data), filtrele=True)
        return olay.action

    def _kabul_edilir_mi(self, yol: str) -> bool:
        return Path(yol).suffix.lower().lstrip(".") in self.bilgi["kabul"]

    def yollari_ekle(self, yollar, filtrele: bool) -> int:
        """Dosya/klasör yollarını listeye ekler (klasörler alt klasörleriyle
        taranır). filtrele=True ise bu sekmenin desteklemediği dosyalar atlanır."""
        mevcut = {os.path.normcase(os.path.normpath(y)) for y in self.liste.get(0, tk.END)}
        eklenen = atlanan = 0
        adaylar = []
        for yol in yollar:
            if os.path.isdir(yol):
                for kok, _klasorler, dosyalar in os.walk(yol):
                    adaylar += [os.path.join(kok, d) for d in sorted(dosyalar)
                                if self._kabul_edilir_mi(d)]
                    if len(adaylar) > _AZAMI_KLASOR_DOSYASI:
                        break
            elif filtrele and not self._kabul_edilir_mi(yol):
                atlanan += 1
            else:
                adaylar.append(yol)
        for yol in adaylar[:_AZAMI_KLASOR_DOSYASI]:
            yol = os.path.normpath(yol)
            anahtar = os.path.normcase(yol)
            if anahtar not in mevcut:
                mevcut.add(anahtar)
                self.liste.insert(tk.END, yol)
                eklenen += 1
        if yollar:
            mesaj = f"{eklenen} dosya eklendi."
            if atlanan:
                mesaj += f" ({atlanan} dosya bu sekmede desteklenmediği için atlandı.)"
            self.durum.set(mesaj)
        return eklenen

    def _dosya_ekle(self):
        dosyalar = filedialog.askopenfilenames(
            title="Dosya seç",
            filetypes=self.bilgi["uzantilar"] + [("Tüm dosyalar", "*.*")])
        self.yollari_ekle(dosyalar, filtrele=False)

    def _klasor_ekle(self):
        klasor = filedialog.askdirectory(title="Eklenecek klasörü seç")
        if klasor:
            self.yollari_ekle([klasor], filtrele=True)

    def _secileni_kaldir(self):
        for indeks in reversed(self.liste.curselection()):
            self.liste.delete(indeks)

    def _klasor_sec(self):
        klasor = filedialog.askdirectory(title="Çıktı klasörü seç")
        if klasor:
            self.cikti_yolu.set(os.path.normpath(klasor))

    # ------------------------------------------------ dönüştürme

    def _donusturmeyi_baslat(self):
        dosyalar = list(self.liste.get(0, tk.END))
        if not dosyalar:
            messagebox.showwarning("Dönüştürücü", "Önce dosya ekleyin.")
            return
        hedef = self.format_kutusu.get()
        cikti = self.cikti_yolu.get()
        if cikti == _VARSAYILAN_CIKTI:
            cikti = None

        self.calisiyor = True
        self.iptal.clear()
        self._dugmeleri_ayarla()
        self.ac_dugmesi.grid_remove()
        self.ilerleme.configure(maximum=len(dosyalar), value=0)
        self.durum.set("Dönüştürülüyor...")

        threading.Thread(
            target=self._arka_planda_donustur,
            args=(dosyalar, hedef, cikti),
            daemon=True,
        ).start()
        self.after(100, self._kuyrugu_isle)

    def _iptal_et(self):
        self.iptal.set()
        self.iptal_dugmesi.configure(state="disabled")
        self.durum.set("İptal ediliyor...")

    def _dugmeleri_ayarla(self):
        durum = "disabled" if self.calisiyor else "normal"
        self.donustur_dugmesi.configure(state=durum)
        for dugme in self.liste_dugmeleri:
            dugme.configure(state=durum)
        self.iptal_dugmesi.configure(state="normal" if self.calisiyor else "disabled")

    def _arka_planda_donustur(self, dosyalar, hedef, cikti):
        hatalar = []
        basarili = 0
        son_cikti = None
        iptal_edildi = False
        try:
            for i, dosya in enumerate(dosyalar, start=1):
                if self.iptal.is_set():
                    iptal_edildi = True
                    break
                self.kuyruk.put(("dosya", i, len(dosyalar), Path(dosya).name))
                try:
                    son_cikti = self.bilgi["modul"].donustur(
                        dosya, hedef, cikti, iptal=self.iptal)
                    basarili += 1
                except IptalEdildi:
                    iptal_edildi = True
                    break
                except Exception as h:  # tek dosyanın hatası diğerlerini durdurmasın
                    hatalar.append(f"{Path(dosya).name}: {hata_metni(h)}")
                self.kuyruk.put(("ilerleme", i))
        finally:
            self.kuyruk.put(("bitti", basarili, hatalar, iptal_edildi, son_cikti))

    def _kuyrugu_isle(self):
        try:
            while True:
                mesaj = self.kuyruk.get_nowait()
                if mesaj[0] == "dosya":
                    _, sira, toplam, ad = mesaj
                    if not self.iptal.is_set():
                        self.durum.set(f"Dönüştürülüyor ({sira}/{toplam}): {ad}")
                elif mesaj[0] == "ilerleme":
                    self.ilerleme.configure(value=mesaj[1])
                elif mesaj[0] == "bitti":
                    self._bitir(*mesaj[1:])
                    return
        except queue.Empty:
            pass
        self.after(100, self._kuyrugu_isle)

    def _bitir(self, basarili, hatalar, iptal_edildi, son_cikti):
        self.calisiyor = False
        self._dugmeleri_ayarla()
        self.son_cikti = son_cikti
        if son_cikti:
            self.ac_dugmesi.grid(row=0, column=1, sticky="e", padx=(8, 0))
        if iptal_edildi:
            self.durum.set(f"⏹ İptal edildi — {basarili} dosya dönüştürüldü.")
        elif hatalar:
            self.durum.set(f"✅ {basarili} dosya dönüştürüldü, ❌ {len(hatalar)} hata.")
        else:
            self.durum.set(f"✅ {basarili} dosya başarıyla dönüştürüldü.")
        if hatalar:
            metin = "\n\n".join(hatalar[:10])
            if len(hatalar) > 10:
                metin += f"\n\n... ve {len(hatalar) - 10} hata daha."
            messagebox.showerror("Dönüştürme hataları", metin)


def _pencere_olustur() -> tuple[tk.Tk, bool]:
    """Sürükle-bırak destekli pencere oluşturmayı dener, olmazsa düz Tk."""
    if TkinterDnD is not None:
        try:
            return TkinterDnD.Tk(), True
        except Exception:
            pass  # tkdnd yerel kütüphanesi yüklenemedi
    return tk.Tk(), False


def main():
    pencere, surukle_birak = _pencere_olustur()
    pencere.title("Dönüştürücü — Belge • Görsel • Ses • Video")
    pencere.geometry("720x640")
    pencere.minsize(660, 580)
    try:
        pencere.iconbitmap(str(_kaynak_yolu("icon.ico")))
    except Exception:
        pass  # simge bulunamazsa sessizce yok say, uygulamayı engellemesin

    tema = _tema_yukle()
    sv_ttk.set_theme(tema)

    ust_bar = ttk.Frame(pencere, padding=(20, 16, 20, 0))
    ust_bar.pack(fill=tk.X)
    ttk.Label(ust_bar, text="🔄 Dönüştürücü", font="SunValleySubtitleFont").pack(
        side=tk.LEFT)

    defter = ttk.Notebook(pencere)
    defter.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)
    sekmeler = [DonusturmeSekmesi(defter, bilgi, surukle_birak) for bilgi in SEKMELER]
    for sekme, bilgi in zip(sekmeler, SEKMELER):
        defter.add(sekme, text=bilgi["ad"])

    def _tema_degistir():
        sv_ttk.toggle_theme(pencere)
        yeni = sv_ttk.get_theme(pencere)
        for sekme in sekmeler:
            sekme.liste_temasini_uygula(yeni)
        tema_dugmesi.configure(text=_tema_dugme_metni(yeni))
        _tema_kaydet(yeni)

    tema_dugmesi = ttk.Button(
        ust_bar, text=_tema_dugme_metni(tema), command=_tema_degistir)
    tema_dugmesi.pack(side=tk.RIGHT)

    def _kapat():
        calisanlar = [s for s in sekmeler if s.calisiyor]
        if calisanlar:
            if not messagebox.askyesno(
                    "Dönüştürücü",
                    "Dönüştürme sürüyor. Yine de çıkmak istiyor musunuz?\n\n"
                    "Yarım kalan dosya silinir.", icon="warning"):
                return
            for sekme in calisanlar:
                sekme.iptal.set()
            # İş parçacıklarının yarım dosyayı silip Word/ffmpeg'i kapatması
            # için kısa süre bekle (en fazla ~5 sn), sonra pencereyi kapat.
            _bekle_ve_kapat(50)
        else:
            pencere.destroy()

    def _bekle_ve_kapat(kalan: int):
        if kalan <= 0 or not any(s.calisiyor for s in sekmeler):
            pencere.destroy()
        else:
            pencere.after(100, _bekle_ve_kapat, kalan - 1)

    pencere.protocol("WM_DELETE_WINDOW", _kapat)
    pencere.mainloop()


if __name__ == "__main__":
    main()
