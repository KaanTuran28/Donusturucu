# Dönüştürücü

Belge, görsel, ses ve video dosyalarını kendi aralarında dönüştüren
profesyonel masaüstü programı (Windows, Türkçe arayüz, açık/koyu tema).

## Kurulum

Gereken: **Windows 10/11** ve **Python 3.10 veya üstü**
([python.org](https://www.python.org/downloads/) — kurulumda
**"Add python.exe to PATH"** kutusunu işaretleyin).

1. Projeyi indirin:
   ```
   git clone https://github.com/KaanTuran28/Donusturucu.git
   cd Donusturucu
   ```
   (Git yoksa GitHub sayfasında **Code → Download ZIP** ile indirip klasöre
   çıkarın ve komut satırını o klasörde açın.)
2. Kütüphaneleri kurun:
   ```
   python -m pip install -r requirements.txt
   ```
3. Çalıştırın: **Başlat.bat** dosyasına çift tıklayın (ya da `python donusturucu.py`).

ffmpeg ayrıca kurulmaz (kütüphaneyle gömülü gelir). Microsoft Word isteğe
bağlıdır: kuruluysa belgeler Word'ün motoruyla birebir dönüştürülür, yoksa
dahili motor kullanılır.

### İsteğe bağlı: Python gerektirmeyen tek klasörlük .exe

```
python -m pip install -r requirements-dev.txt
python -m PyInstaller Donusturucu.spec --distpath build_cikti/dist --workpath build_cikti/build --noconfirm
```

Oluşan `build_cikti\dist\Donusturucu\` klasörü tek başına çalışır; örneğin
`%LOCALAPPDATA%\Programs\Dönüştürücü\` altına kopyalayıp içindeki
`Donusturucu.exe` için masaüstü kısayolu oluşturabilirsiniz. Güncellerken
eski dosyalar kalmasın diye klasörü önce silin.

## Desteklenen formatlar

| Tür | Formatlar |
|---|---|
| 📄 Belgeler | PDF, DOCX, TXT, HTML, MD, RTF, ODT (+ Word kuruluysa kaynak olarak eski **DOC**) |
| 🖼️ Görseller | PNG, JPG, BMP, GIF, WEBP, TIFF, ICO, AVIF, HEIC + görselden PDF |
| 🎵 Sesler | MP3, WAV, OGG, FLAC, M4A, AAC, WMA, OPUS, AIFF + videodan ses çıkarma |
| 🎬 Videolar | MP4, MKV, MOV, AVI, WEBM, GIF + videodan MP3/WAV çıkarma |

## Dönüşüm kalitesi (belgeler)

Program en yüksek kaliteli yolu otomatik seçer:

1. **Microsoft Word kuruluysa** (bu bilgisayarda kurulu ✓): DOCX, DOC, RTF,
   ODT ve HTML kaynakları Word'ün kendi motoruyla dönüştürülür — tablolar,
   şekiller, resimler ve stiller **birebir** korunur. HTML/MD'deki resimler
   çıktıya gömülür (bağlantı olarak bırakılmaz).
2. **PDF → DOCX**: `pdf2docx` ile sayfa düzeni (tablolar, sütunlar, resimler)
   korunarak dönüştürülür.
3. Diğer yollar: dahili blok motoru — başlıklar, kalın/italik vurgular,
   listeler, tablolar (metin olarak), köprü metinleri ve gömülü resimler
   korunur.

## Kullanım

1. İlgili sekmeyi seçin (Belgeler / Görseller / Sesler / Videolar).
2. Dosyaları listeye **sürükleyip bırakın**, ya da **➕ Dosya Ekle** /
   **📁 Klasör Ekle** düğmelerini kullanın (klasörler alt klasörleriyle
   taranır; sekmeye uymayan dosyalar atlanır).
3. **Hedef format**ı seçin.
4. İsterseniz **çıktı klasörü** seçin (seçmezseniz dönüştürülen dosya
   kaynak dosyanın yanına kaydedilir; **↺** ile bu varsayılana dönülür).
5. **🔄 DÖNÜŞTÜR** düğmesine basın. Uzun işlemleri **⏹ İptal** ile
   durdurabilirsiniz; bittiğinde **📂 Çıktıyı Göster** dosyayı Gezgin'de açar.

İpuçları:
- Sesler sekmesine **video dosyası** ekleyip MP3 seçerseniz videonun sesi çıkarılır.
- Program **hiçbir dosyanın üzerine yazmaz**: hedefte aynı adlı bir dosya
  varsa çıktı `ad_donusturuldu`, `ad_donusturuldu_2` ... adıyla kaydedilir.
- Listede seçili dosyaları **Delete** tuşuyla kaldırabilirsiniz.
- Sağ üstteki **🌙 Koyu Tema / ☀️ Açık Tema** düğmesiyle görünümü
  değiştirebilirsiniz; seçiminiz hatırlanır.

## Güvenlik

- Güvenilmeyen HTML/MD dosyalarındaki resim adresleri denetlenir: ağ
  paylaşımı (UNC, ör. `\\sunucu\x.png`) ve URL'ler reddedilir, yalnızca
  gerçekten resim olan yerel dosyalar gömülür. Ağ adresine başvuran HTML'ler
  Word'e verilmez (Windows kimlik bilgisi özetinin sızmasını önlemek için).
- Belgelerden çıkarılan resim adları temizlenir (klasör dışına dosya yazılamaz).
- Word ile açılan belgelerde makrolar zorla kapalıdır; parolalı belgeler Word'e
  verilmeden tespit edilir; Word takılırsa iletişim kutuları otomatik kapatılır
  ve süre sınırında süreç sonlandırılır.
- ffmpeg yalnızca yerel dosya okuyabilir (`-protocol_whitelist file`).
- Bağımlılıkların alt sürümleri bilinen açıkları kapatan sürümlere
  sabitlenmiştir; denetlemek için: `python -m pip_audit`.

## Testler

```
pip install -r requirements-dev.txt
python -m pytest
```

~450 test (~4,5 dk): tüm format çiftleri (belgelerde Word ile ve Word olmadan), görsel
renk kipleri/EXIF/animasyon, ses/video (tek sayılı boyutlar, çok kanallı ses,
iptal), güvenlik gerileme testleri ve arayüz duman testi. Word yoksa Word
testleri otomatik atlanır.

## Gereksinimler

- Python 3.10+ (3.14 ile test edildi)
- Kütüphaneler: `pip install -r requirements.txt`
- Ses/video için ffmpeg kurmaya gerek yok (`imageio-ffmpeg` gömülü gelir).
- Microsoft Word isteğe bağlıdır; yoksa dahili motor devreye girer
  (yalnızca eski .DOC dosyaları için Word gerekir).

## Notlar

- Taranmış (görüntüden oluşan) PDF'lerden metin çıkarılamaz.
- Parola ile korunan belgeler/PDF'ler dönüştürülemez (yalnızca izin/sahip
  parolalı PDF'ler otomatik açılır).
- Saydam görseller JPG/BMP/PDF'e dönüştürülürken beyaz arka plan uygulanır.
- Telefon fotoğraflarının EXIF yönü uygulanır; renk profili (ICC) korunur.
- Animasyonlu GIF/WEBP → GIF/WEBP/PNG dönüşümlerinde animasyon, çok sayfalı
  TIFF → PDF/TIFF dönüşümlerinde tüm sayfalar korunur.
- ICO çıktısı 16×16–256×256 piksel arasıdır (format sınırı); 16px'in
  altındaki kaynaklar otomatik büyütülür.
- Video dönüşümleri dosya boyutuna göre zaman alabilir.
