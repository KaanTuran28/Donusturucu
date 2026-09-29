"""Arayüz duman testi: pencere kurulur, dosya eklenir, dönüştürülür, iptal edilir."""

import time
from pathlib import Path

import pytest
from conftest import ffmpeg
from PIL import Image

tk = pytest.importorskip("tkinter")
import donusturucu as uyg  # noqa: E402


@pytest.fixture
def pencere():
    try:
        kok, dnd = uyg._pencere_olustur()
    except tk.TclError as hata:
        pytest.skip(f"ekran yok: {hata}")
    uyg.sv_ttk.set_theme("light")
    defter = uyg.ttk.Notebook(kok)
    defter.pack()
    sekmeler = [uyg.DonusturmeSekmesi(defter, b, dnd) for b in uyg.SEKMELER]
    for s, b in zip(sekmeler, uyg.SEKMELER):
        defter.add(s, text=b["ad"])
    kok.update()
    yield kok, sekmeler
    kok.destroy()


def _bekle(kok, sekme, sure=60):
    son = time.time() + sure
    while sekme.calisiyor and time.time() < son:
        kok.update()
        time.sleep(0.03)
    assert not sekme.calisiyor


def test_klasor_ekleme_filtre_ve_donusum(pencere, tmp_path):
    kok, sekmeler = pencere
    gorsel = sekmeler[1]
    (tmp_path / "alt").mkdir()
    for i in range(3):
        Image.new("RGB", (20, 20), "red").save(tmp_path / "alt" / f"r{i}.png")
    (tmp_path / "alt" / "belge.txt").write_text("x", encoding="utf-8")

    assert gorsel.yollari_ekle([str(tmp_path / "alt"), str(tmp_path / "x.txt")], True) == 3
    assert "atlandı" in gorsel.durum.get()
    assert gorsel.yollari_ekle([str(tmp_path / "alt")], True) == 0  # çift eklenmez

    gorsel.format_kutusu.set("webp")
    gorsel._donusturmeyi_baslat()
    kok.update()
    assert str(gorsel.donustur_dugmesi["state"]) == "disabled"
    _bekle(kok, gorsel)
    assert "3 dosya başarıyla" in gorsel.durum.get()
    assert gorsel.son_cikti and Path(gorsel.son_cikti).exists()
    assert str(gorsel.donustur_dugmesi["state"]) == "normal"


def test_iptal_dugmesi(pencere, tmp_path):
    kok, sekmeler = pencere
    video = sekmeler[3]
    ffmpeg("-f", "lavfi", "-i", "testsrc=size=1280x720:rate=30:duration=60",
           "-c:v", "libx264", "-preset", "ultrafast", tmp_path / "uzun.mp4")
    video.yollari_ekle([str(tmp_path / "uzun.mp4")], True)
    video.format_kutusu.set("webm")
    video._donusturmeyi_baslat()
    for _ in range(20):
        kok.update()
        time.sleep(0.05)
    video._iptal_et()
    _bekle(kok, video, 10)
    assert "İptal edildi" in video.durum.get()
    assert not list(tmp_path.glob("uzun*.webm"))


def test_tema_listeyi_gunceller(pencere):
    _, sekmeler = pencere
    for sekme in sekmeler:
        sekme.liste_temasini_uygula("dark")
        assert sekme.liste["background"] == uyg._LISTE_RENKLERI["dark"]["arka"]
