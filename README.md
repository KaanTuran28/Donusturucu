# Dönüştürücü — File Converter

![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-0078D6)
![License](https://img.shields.io/badge/license-MIT-green)

<p align="center"><b><a href="#english">English</a></b> · <b><a href="#türkçe">Türkçe</a></b></p>

---

## English

A professional Windows desktop app that converts documents, images, audio and video between formats. Turkish interface, light/dark theme.

### Installation

Requires **Windows 10/11** and **Python 3.10+** ([python.org](https://www.python.org/downloads/) — tick **"Add python.exe to PATH"** during install).

```bash
git clone https://github.com/KaanTuran28/Donusturucu.git
cd Donusturucu
python -m pip install -r requirements.txt
```

Run it by double-clicking **Başlat.bat**, or with `python donusturucu.py`.

ffmpeg does not need separate installation (it ships embedded). Microsoft Word is optional: if installed, documents are converted with Word's own engine; otherwise the built-in engine is used.

**Optional — a standalone .exe that needs no Python:**

```bash
python -m pip install -r requirements-dev.txt
python -m PyInstaller Donusturucu.spec --distpath build_cikti/dist --workpath build_cikti/build --noconfirm
```

The resulting `build_cikti\dist\Donusturucu\` folder runs on its own.

### Supported formats

| Type | Formats |
|---|---|
| 📄 Documents | PDF, DOCX, TXT, HTML, MD, RTF, ODT (+ legacy **DOC** as a source if Word is installed) |
| 🖼️ Images | PNG, JPG, BMP, GIF, WEBP, TIFF, ICO, AVIF, HEIC + image→PDF |
| 🎵 Audio | MP3, WAV, OGG, FLAC, M4A, AAC, WMA, OPUS, AIFF + extract audio from video |
| 🎬 Video | MP4, MKV, MOV, AVI, WEBM, GIF + extract MP3/WAV from video |

### Conversion quality (documents)

The app picks the highest-quality path automatically:

1. **If Microsoft Word is installed:** DOCX, DOC, RTF, ODT and HTML sources are converted with Word's own engine — tables, shapes, images and styles are preserved exactly. Images in HTML/MD are embedded in the output.
2. **PDF → DOCX:** converted with `pdf2docx`, preserving page layout (tables, columns, images).
3. **Other paths:** the built-in block engine preserves headings, bold/italic, lists, tables (as text), hyperlinks and embedded images.

### Usage

1. Pick a tab (Documents / Images / Audio / Video).
2. **Drag and drop** files into the list, or use **➕ Add File** / **📁 Add Folder** (folders are scanned recursively; files that don't match the tab are skipped).
3. Choose the **target format**.
4. Optionally choose an **output folder** (otherwise the converted file is saved next to the source; **↺** restores the default).
5. Press **🔄 CONVERT**. Long jobs can be stopped with **⏹ Cancel**; **📂 Show Output** opens the result in Explorer.

Tips: add a **video** on the Audio tab and pick MP3 to extract its audio. The app **never overwrites** anything — a name clash is saved as `name_donusturuldu`, `name_donusturuldu_2`, … Remove selected files with **Delete**. The **🌙/☀️** button toggles the theme and your choice is remembered.

### Security

- Image paths in untrusted HTML/MD are checked: network shares (UNC, e.g. `\\server\x.png`) and URLs are rejected, and only real local image files are embedded. HTML referencing a network address is not handed to Word (to prevent leaking a Windows credential hash).
- Image names extracted from documents are sanitised (no writing outside the folder).
- Macros are forced off for Word-opened documents; password-protected documents are detected before being handed to Word; if Word hangs, dialogs are auto-closed and the process is terminated on timeout.
- ffmpeg can only read local files (`-protocol_whitelist file`).
- Dependency lower bounds are pinned to versions that close known vulnerabilities; audit with `python -m pip_audit`.

### Tests

```bash
pip install -r requirements-dev.txt
python -m pytest
```

~450 tests (~4.5 min): all format pairs (documents with and without Word), image colour modes/EXIF/animation, audio/video (odd dimensions, multi-channel audio, cancellation), security regression tests and a UI smoke test. Word tests are skipped automatically when Word is absent.

### Notes

- Text cannot be extracted from scanned (image-only) PDFs.
- Password-protected documents/PDFs cannot be converted (only permission/owner-password PDFs open automatically).
- A white background is applied when transparent images are converted to JPG/BMP/PDF.
- EXIF orientation of phone photos is applied; the ICC colour profile is preserved.
- Animation is preserved in animated GIF/WEBP conversions, and all pages in multi-page TIFF → PDF/TIFF.
- ICO output is 16×16–256×256 px (format limit); smaller sources are upscaled.
- Video conversions can take time depending on file size.

### License

The source code is distributed under the [MIT license](LICENSE). The program uses third-party libraries under their own licenses (installed via `requirements.txt`). In particular **pdf2docx** (used for PDF → DOCX) builds on **PyMuPDF** (AGPL-3.0), and the **ffmpeg** shipped via `imageio-ffmpeg` is GPL-licensed. If you distribute a prebuilt `.exe` bundled with these libraries, you must also comply with those licenses.

---

## Türkçe

Belge, görsel, ses ve video dosyalarını kendi aralarında dönüştüren profesyonel masaüstü programı (Windows, Türkçe arayüz, açık/koyu tema).

### Kurulum

Gereken: **Windows 10/11** ve **Python 3.10 veya üstü** ([python.org](https://www.python.org/downloads/) — kurulumda **"Add python.exe to PATH"** kutusunu işaretleyin).

```bash
git clone https://github.com/KaanTuran28/Donusturucu.git
cd Donusturucu
python -m pip install -r requirements.txt
```

Çalıştırmak için **Başlat.bat** dosyasına çift tıklayın (ya da `python donusturucu.py`).

ffmpeg ayrıca kurulmaz (kütüphaneyle gömülü gelir). Microsoft Word isteğe bağlıdır: kuruluysa belgeler Word'ün motoruyla birebir dönüştürülür, yoksa dahili motor kullanılır.

**İsteğe bağlı — Python gerektirmeyen tek klasörlük .exe:**

```bash
python -m pip install -r requirements-dev.txt
python -m PyInstaller Donusturucu.spec --distpath build_cikti/dist --workpath build_cikti/build --noconfirm
```

Oluşan `build_cikti\dist\Donusturucu\` klasörü tek başına çalışır; güncellerken eski dosyalar kalmasın diye klasörü önce silin.

### Desteklenen formatlar

| Tür | Formatlar |
|---|---|
| 📄 Belgeler | PDF, DOCX, TXT, HTML, MD, RTF, ODT (+ Word kuruluysa kaynak olarak eski **DOC**) |
| 🖼️ Görseller | PNG, JPG, BMP, GIF, WEBP, TIFF, ICO, AVIF, HEIC + görselden PDF |
| 🎵 Sesler | MP3, WAV, OGG, FLAC, M4A, AAC, WMA, OPUS, AIFF + videodan ses çıkarma |
| 🎬 Videolar | MP4, MKV, MOV, AVI, WEBM, GIF + videodan MP3/WAV çıkarma |

### Dönüşüm kalitesi (belgeler)

Program en yüksek kaliteli yolu otomatik seçer:

1. **Microsoft Word kuruluysa:** DOCX, DOC, RTF, ODT ve HTML kaynakları Word'ün kendi motoruyla dönüştürülür — tablolar, şekiller, resimler ve stiller **birebir** korunur. HTML/MD'deki resimler çıktıya gömülür.
2. **PDF → DOCX:** `pdf2docx` ile sayfa düzeni (tablolar, sütunlar, resimler) korunarak dönüştürülür.
3. **Diğer yollar:** dahili blok motoru — başlıklar, kalın/italik vurgular, listeler, tablolar (metin olarak), köprü metinleri ve gömülü resimler korunur.

### Kullanım

1. İlgili sekmeyi seçin (Belgeler / Görseller / Sesler / Videolar).
2. Dosyaları listeye **sürükleyip bırakın**, ya da **➕ Dosya Ekle** / **📁 Klasör Ekle** düğmelerini kullanın (klasörler alt klasörleriyle taranır; sekmeye uymayan dosyalar atlanır).
3. **Hedef format**ı seçin.
4. İsterseniz **çıktı klasörü** seçin (seçmezseniz dönüştürülen dosya kaynak dosyanın yanına kaydedilir; **↺** ile varsayılana dönülür).
5. **🔄 DÖNÜŞTÜR** düğmesine basın. Uzun işlemleri **⏹ İptal** ile durdurabilirsiniz; **📂 Çıktıyı Göster** dosyayı Gezgin'de açar.

İpuçları: Sesler sekmesine **video dosyası** ekleyip MP3 seçerseniz videonun sesi çıkarılır. Program **hiçbir dosyanın üzerine yazmaz**: aynı adlı dosya varsa çıktı `ad_donusturuldu`, `ad_donusturuldu_2` ... adıyla kaydedilir. Seçili dosyaları **Delete** tuşuyla kaldırabilirsiniz. Sağ üstteki **🌙/☀️** düğmesi temayı değiştirir ve seçiminiz hatırlanır.

### Güvenlik

- Güvenilmeyen HTML/MD'deki resim adresleri denetlenir: ağ paylaşımı (UNC, ör. `\\sunucu\x.png`) ve URL'ler reddedilir, yalnızca gerçekten resim olan yerel dosyalar gömülür. Ağ adresine başvuran HTML'ler Word'e verilmez (Windows kimlik bilgisi özetinin sızmasını önlemek için).
- Belgelerden çıkarılan resim adları temizlenir (klasör dışına dosya yazılamaz).
- Word ile açılan belgelerde makrolar zorla kapalıdır; parolalı belgeler Word'e verilmeden tespit edilir; Word takılırsa iletişim kutuları otomatik kapatılır ve süre sınırında süreç sonlandırılır.
- ffmpeg yalnızca yerel dosya okuyabilir (`-protocol_whitelist file`).
- Bağımlılıkların alt sürümleri bilinen açıkları kapatan sürümlere sabitlenmiştir; denetlemek için: `python -m pip_audit`.

### Testler

```bash
pip install -r requirements-dev.txt
python -m pytest
```

~450 test (~4,5 dk): tüm format çiftleri (belgelerde Word ile ve Word olmadan), görsel renk kipleri/EXIF/animasyon, ses/video (tek sayılı boyutlar, çok kanallı ses, iptal), güvenlik gerileme testleri ve arayüz duman testi. Word yoksa Word testleri otomatik atlanır.

### Notlar

- Taranmış (görüntüden oluşan) PDF'lerden metin çıkarılamaz.
- Parola ile korunan belgeler/PDF'ler dönüştürülemez (yalnızca izin/sahip parolalı PDF'ler otomatik açılır).
- Saydam görseller JPG/BMP/PDF'e dönüştürülürken beyaz arka plan uygulanır.
- Telefon fotoğraflarının EXIF yönü uygulanır; renk profili (ICC) korunur.
- Animasyonlu GIF/WEBP ve çok sayfalı TIFF → PDF/TIFF dönüşümlerinde tüm içerik korunur.
- ICO çıktısı 16×16–256×256 piksel arasıdır (format sınırı); daha küçük kaynaklar otomatik büyütülür.
- Video dönüşümleri dosya boyutuna göre zaman alabilir.

### Lisans

Kaynak kod [MIT lisansı](LICENSE) ile dağıtılır. Program, kendi lisanslarına tabi üçüncü taraf kütüphaneler kullanır (`requirements.txt` ile kurulur). Özellikle PDF → DOCX için kullanılan **pdf2docx**, **PyMuPDF**'e (AGPL-3.0) dayanır; `imageio-ffmpeg` ile gelen **ffmpeg** GPL lisanslıdır. Programı bu kütüphanelerle paketlenmiş hazır bir .exe olarak dağıtacaksanız bu lisansların koşullarına da uymanız gerekir.
