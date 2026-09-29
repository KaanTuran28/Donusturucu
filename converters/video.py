"""Video dönüştürücü - gömülü ffmpeg (imageio-ffmpeg) tabanlı.

Desteklenen hedefler: MP4, MKV, MOV, AVI, WEBM, GIF
Ayrıca videodan ses çıkarma: MP3, WAV
"""

import shutil
from pathlib import Path

from . import ortak

FORMATLAR = ["mp4", "mkv", "mov", "avi", "webm", "gif", "mp3", "wav"]

# yuv420p renk alt örneklemesi genişlik/yüksekliğin çift sayı olmasını
# gerektirir; tek sayılı boyutlu kaynaklar (ör. 101x51 bir GIF ya da ekran
# kaydı) bu filtre olmadan "width not divisible by 2" hatasıyla çöküyordu.
_CIFT_BOYUT = ["-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2"]
_H264 = ["-c:v", "libx264", "-preset", "fast", "-crf", "23",
         "-pix_fmt", "yuv420p", *_CIFT_BOYUT]

# Kaliteli GIF: önce videoya özel 256 renklik palet üretilir, sonra o paletle
# kodlanır (varsayılan genel palete göre çok daha az bantlaşma/renk kaybı).
# Küçük videolar büyütülmez: genişlik en fazla 480px.
_GIF_FILTRE = ("fps=12,scale='min(480,iw)':-1:flags=lanczos,"
               "split[a][b];[a]palettegen=stats_mode=diff[p];"
               "[b][p]paletteuse=dither=bayer:bayer_scale=5")

_AYARLAR = {
    # +faststart: dosya indirilirken/akışta hemen oynatılabilsin
    "mp4": [*_H264, "-c:a", "aac", "-b:a", "160k", "-movflags", "+faststart"],
    "mkv": [*_H264, "-c:a", "aac", "-b:a", "160k"],
    "mov": [*_H264, "-c:a", "aac", "-b:a", "160k", "-movflags", "+faststart"],
    "avi": ["-c:v", "mpeg4", "-q:v", "4", *_CIFT_BOYUT,
            "-c:a", "libmp3lame", "-b:a", "160k"],
    # -pix_fmt yuv420p şart: GIF gibi saydamlık taşıyan kaynaklarda libvpx
    # "auto_alt_ref" ile saydam kareleri kodlayamayıp hata veriyor.
    "webm": ["-c:v", "libvpx", "-pix_fmt", "yuv420p", *_CIFT_BOYUT,
             "-b:v", "1M", "-c:a", "libvorbis"],
    "gif": ["-vf", _GIF_FILTRE, "-loop", "0", "-an"],
    # videodan ses çıkarma
    "mp3": ["-vn", "-c:a", "libmp3lame", "-b:a", "192k"],
    "wav": ["-vn", "-c:a", "pcm_s16le"],
}


def donustur(kaynak: str, hedef_format: str, cikti_klasoru: str | None = None,
             iptal=None) -> str:
    """Bir video dosyasını hedef formata dönüştürür, çıktı yolunu döndürür."""
    kaynak_yol = Path(kaynak)
    kaynak_format = kaynak_yol.suffix.lower().lstrip(".")
    hedef_format = hedef_format.lower().lstrip(".")
    if hedef_format not in _AYARLAR:
        raise ValueError(f"Desteklenmeyen video formatı: {hedef_format}")
    ortak.kaynak_denetle(kaynak_yol)

    akis = "ses" if hedef_format in ("mp3", "wav") else "video/ses"
    with ortak.cikti_dosyasi(kaynak_yol, hedef_format, cikti_klasoru) as cikti:
        if kaynak_format == hedef_format:
            # Aynı format: yeniden kodlamadan (kalite kaybı olmadan) doğrudan kopyala
            shutil.copy2(kaynak_yol, cikti)
        else:
            ortak.ffmpeg_calistir(kaynak_yol, _AYARLAR[hedef_format], cikti,
                                  iptal, akis=akis)
    return str(cikti)
