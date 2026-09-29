"""Ses ve video dönüşümleri: tüm format çiftleri + iptal ve hata mesajları."""

import threading
import time
from pathlib import Path

import pytest
from conftest import ffmpeg, okunabilir_medya

from converters import ortak, ses, video

SES_KAYNAKLARI = ["mp3", "wav", "ogg", "flac", "m4a", "aac", "wma", "opus", "aiff"]
VIDEO_KAYNAKLARI = ["mp4", "mkv", "mov", "avi", "webm", "gif"]


@pytest.fixture(scope="module")
def medya(tmp_path_factory) -> dict[str, Path]:
    kl = tmp_path_factory.mktemp("medya ornek ş")  # boşluk + Türkçe karakter
    ornek = {}
    wav = kl / "ses.wav"
    ffmpeg("-f", "lavfi", "-i", "sine=frequency=440:duration=1", wav)
    for f in SES_KAYNAKLARI:
        if f == "wav":
            ornek[f] = wav
            continue
        ornek[f] = Path(ses.donustur(str(wav), f, str(kl / "ses")))
    # tek sayılı boyutlu video (eskiden H.264 hedeflerinde çöküyordu)
    ana = kl / "video.mkv"
    ffmpeg("-f", "lavfi", "-i", "testsrc=size=161x121:rate=15:duration=1",
           "-f", "lavfi", "-i", "sine=duration=1", "-shortest",
           "-c:v", "libx264", "-pix_fmt", "yuv444p", "-c:a", "aac", ana)
    for f in VIDEO_KAYNAKLARI:
        ornek["v_" + f] = ana if f == "mkv" else Path(
            video.donustur(str(ana), f, str(kl / "video")))
    ornek["kanal51"] = kl / "kanal51.wav"
    ffmpeg("-f", "lavfi", "-i", "sine=duration=1", "-af",
           "pan=5.1|c0=c0|c1=c0|c2=c0|c3=c0|c4=c0|c5=c0", ornek["kanal51"])
    return ornek


@pytest.mark.parametrize("hedef", ses.FORMATLAR)
@pytest.mark.parametrize("kaynak", SES_KAYNAKLARI + ["v_mp4"])
def test_ses_format_ciftleri(medya, tmp_path, kaynak, hedef):
    cikti = Path(ses.donustur(str(medya[kaynak]), hedef, str(tmp_path)))
    assert cikti.suffix == "." + hedef
    assert okunabilir_medya(cikti)


@pytest.mark.parametrize("hedef", video.FORMATLAR)
@pytest.mark.parametrize("kaynak", VIDEO_KAYNAKLARI)
def test_video_format_ciftleri(medya, tmp_path, kaynak, hedef):
    if kaynak == "gif" and hedef in ("mp3", "wav"):
        pytest.skip("GIF'te ses yok (ayrı testte denetleniyor)")
    cikti = Path(video.donustur(str(medya["v_" + kaynak]), hedef, str(tmp_path)))
    assert okunabilir_medya(cikti)


@pytest.mark.parametrize("hedef", ses.FORMATLAR)
def test_cok_kanalli_ses(medya, tmp_path, hedef):
    assert okunabilir_medya(Path(ses.donustur(str(medya["kanal51"]), hedef, str(tmp_path))))


def test_sessiz_kaynakta_anlasilir_hata(medya, tmp_path):
    with pytest.raises(RuntimeError, match="ses akışı bulunamadı"):
        video.donustur(str(medya["v_gif"]), "mp3", str(tmp_path))
    assert not list(tmp_path.glob("*")), "yarım çıktı silinmeli"


def test_bozuk_dosya_anlasilir_hata(tmp_path):
    (tmp_path / "bozuk.mp4").write_bytes(b"\0" * 2048)
    with pytest.raises(RuntimeError, match="okunamadı"):
        video.donustur(str(tmp_path / "bozuk.mp4"), "webm", str(tmp_path / "o"))


def test_ffmpeg_hata_metni_secimi():
    stderr = (
        "Stream #0:0: Audio: pcm_s16le, 44100 Hz, 5.1\n"
        "[wmav2 @ 0000019a24d4af40] too many channels: got 6, need 2 or fewer\n"
        "[aost#0:0/wmav2 @ 0000019a24d56e40] Error while opening encoder - maybe incorrect\n"
        "[af#0:0 @ 0000019a24d4b540] Task finished with error code: -22 (Invalid argument)\n"
        "[af#0:0 @ 0000019a24d4b540] Terminating thread with return code -22 (Invalid argument)\n"
        "[out#0/asf @ 0000019a24d56d00] Nothing was written into output file\n"
        "size=       0KiB time=N/A bitrate=N/A speed=N/A\nConversion failed!\n")
    assert ortak.ffmpeg_hata_metni(stderr) == "too many channels: got 6, need 2 or fewer"


def test_iptal_ffmpegi_durdurur_ve_yarim_dosyayi_siler(tmp_path):
    uzun = tmp_path / "uzun.mp4"
    ffmpeg("-f", "lavfi", "-i", "testsrc=size=1280x720:rate=30:duration=60",
           "-c:v", "libx264", "-preset", "ultrafast", uzun)
    iptal = threading.Event()
    threading.Timer(0.8, iptal.set).start()
    baslangic = time.monotonic()
    with pytest.raises(ortak.IptalEdildi):
        video.donustur(str(uzun), "webm", str(tmp_path / "o"), iptal=iptal)
    assert time.monotonic() - baslangic < 5
    assert not list((tmp_path / "o").glob("*"))
