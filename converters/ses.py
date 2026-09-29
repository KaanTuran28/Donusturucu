"""Ses dönüştürücü - gömülü ffmpeg (imageio-ffmpeg) tabanlı.

Desteklenen formatlar: MP3, WAV, OGG, FLAC, M4A, AAC, WMA, OPUS, AIFF
Video dosyalarından ses çıkarma da desteklenir (ör. MP4 -> MP3).
"""

import shutil
from pathlib import Path

from . import ortak

FORMATLAR = ["mp3", "wav", "ogg", "flac", "m4a", "aac", "wma", "opus", "aiff"]

# Bazı formatlar için açıkça kodek belirtmek gerekir
_KODEK = {
    "mp3": ["-c:a", "libmp3lame", "-b:a", "192k"],
    "wav": ["-c:a", "pcm_s16le"],
    "ogg": ["-c:a", "libvorbis", "-q:a", "5"],
    "flac": ["-c:a", "flac"],
    "m4a": ["-c:a", "aac", "-b:a", "192k"],
    "aac": ["-c:a", "aac", "-b:a", "192k"],
    # WMA en fazla 2 kanal destekler: 5.1 gibi kaynaklar stereo'ya indirilir
    # (mono kaynak mono kalır). Bu olmadan çok kanallı seste çöküyordu.
    "wma": ["-c:a", "wmav2", "-b:a", "192k",
            "-af", "aformat=channel_layouts=stereo|mono"],
    "opus": ["-c:a", "libopus", "-b:a", "128k"],
    "aiff": ["-c:a", "pcm_s16be"],
}


def donustur(kaynak: str, hedef_format: str, cikti_klasoru: str | None = None,
             iptal=None) -> str:
    """Bir ses dosyasını hedef formata dönüştürür, çıktı yolunu döndürür."""
    kaynak_yol = Path(kaynak)
    kaynak_format = kaynak_yol.suffix.lower().lstrip(".")
    hedef_format = hedef_format.lower().lstrip(".")
    if hedef_format not in _KODEK:
        raise ValueError(f"Desteklenmeyen ses formatı: {hedef_format}")
    ortak.kaynak_denetle(kaynak_yol)

    with ortak.cikti_dosyasi(kaynak_yol, hedef_format, cikti_klasoru) as cikti:
        if kaynak_format == hedef_format:
            # Aynı format: yeniden kodlamadan (kalite kaybı olmadan) doğrudan kopyala
            shutil.copy2(kaynak_yol, cikti)
        else:
            # -vn: video akışlarını atla (kapak resmi vb.)
            ortak.ffmpeg_calistir(kaynak_yol, ["-vn", *_KODEK[hedef_format]],
                                  cikti, iptal, akis="ses")
    return str(cikti)
