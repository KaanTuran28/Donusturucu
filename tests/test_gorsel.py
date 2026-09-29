"""Görsel dönüşümleri: tüm format çiftleri + renk kipi/EXIF/animasyon testleri."""

from pathlib import Path

import pytest
from PIL import Image, ImageSequence

from converters import gorsel

KAYNAK_BICIMLERI = {"png": "PNG", "jpg": "JPEG", "bmp": "BMP", "gif": "GIF",
                    "webp": "WEBP", "tiff": "TIFF", "ico": "ICO", "avif": "AVIF",
                    "heic": "HEIF"}


@pytest.fixture(scope="module")
def gorsel_ornekleri(tmp_path_factory) -> dict[str, Path]:
    kl = tmp_path_factory.mktemp("gorsel_ornek")
    im = Image.new("RGBA", (64, 48), (200, 30, 30, 255))
    im.paste((30, 30, 200, 0), (0, 0, 20, 20))  # saydam köşe
    ornek = {}
    for uzanti, bicim in KAYNAK_BICIMLERI.items():
        yol = kl / f"ornek.{uzanti}"
        kayit = im if bicim in ("PNG", "WEBP", "TIFF", "ICO", "AVIF", "HEIF", "GIF") \
            else im.convert("RGB")
        kayit.save(yol, bicim)
        ornek[uzanti] = yol
    return ornek


def _ac(yol: Path):
    if yol.suffix == ".pdf":
        from pypdf import PdfReader

        assert len(PdfReader(str(yol)).pages) >= 1
        return None
    im = Image.open(yol)
    im.load()
    return im


@pytest.mark.parametrize("hedef", gorsel.FORMATLAR)
@pytest.mark.parametrize("kaynak", list(KAYNAK_BICIMLERI))
def test_tum_format_ciftleri(gorsel_ornekleri, tmp_path, kaynak, hedef):
    cikti = Path(gorsel.donustur(str(gorsel_ornekleri[kaynak]), hedef, str(tmp_path)))
    assert cikti.stat().st_size > 0
    im = _ac(cikti)
    if im is not None:
        assert im.format == gorsel._PIL_FORMAT[hedef]


@pytest.mark.parametrize("hedef", ["png", "webp", "gif", "ico", "avif", "heic", "tiff", "bmp"])
def test_cmyk_kaynak(tmp_path, hedef):
    Image.new("CMYK", (40, 30), (0, 100, 100, 0)).save(tmp_path / "c.jpg")
    im = _ac(Path(gorsel.donustur(str(tmp_path / "c.jpg"), hedef, str(tmp_path / "o"))))
    assert im is not None


def test_16bit_gri_bembeyaz_cikmaz(tmp_path):
    im = Image.new("I;16", (4, 1))
    for x, deger in enumerate((0, 16384, 32768, 65535)):
        im.putpixel((x, 0), deger)
    im.save(tmp_path / "g16.png")
    for hedef in ("jpg", "webp", "bmp"):
        cikti = Image.open(gorsel.donustur(str(tmp_path / "g16.png"), hedef, str(tmp_path / "o")))
        gri = [cikti.convert("L").getpixel((x, 0)) for x in range(4)]
        assert gri[0] < 10 and 50 < gri[1] < 80 and 110 < gri[2] < 145 and gri[3] > 245, gri
    # PNG hedefi 16-bit'i korur
    assert Image.open(gorsel.donustur(str(tmp_path / "g16.png"), "tiff",
                                      str(tmp_path / "o"))).mode.startswith("I;16")


def test_exif_yonu_uygulanir(tmp_path):
    im = Image.new("RGB", (40, 20), "red")
    exif = Image.Exif()
    exif[0x0112] = 6  # 90° döndür
    im.save(tmp_path / "yon.jpg", exif=exif.tobytes())
    for hedef in ("png", "webp", "pdf"):
        cikti = Path(gorsel.donustur(str(tmp_path / "yon.jpg"), hedef, str(tmp_path / "o")))
        if hedef != "pdf":
            assert Image.open(cikti).size == (20, 40)


def test_animasyon_korunur(tmp_path):
    kareler = [Image.new("RGB", (30, 20), r) for r in ("red", "green", "blue")]
    kareler[0].save(tmp_path / "a.gif", save_all=True, append_images=kareler[1:],
                    duration=120, loop=0)
    for hedef in ("webp", "png"):
        cikti = Image.open(gorsel.donustur(str(tmp_path / "a.gif"), hedef, str(tmp_path / "o")))
        assert getattr(cikti, "n_frames", 1) == 3, hedef
    geri = Image.open(gorsel.donustur(str(tmp_path / "o" / "a.webp"), "gif", str(tmp_path / "o2")))
    assert geri.n_frames == 3
    # GIF -> PDF her kareyi sayfa yapmamalı
    from pypdf import PdfReader

    pdf = gorsel.donustur(str(tmp_path / "a.gif"), "pdf", str(tmp_path / "o"))
    assert len(PdfReader(pdf).pages) == 1


def test_cok_sayfali_tiff_pdf(tmp_path):
    from pypdf import PdfReader

    sayfalar = [Image.new("RGB", (100, 140), r) for r in ("white", "gray", "black")]
    sayfalar[0].save(tmp_path / "t.tiff", save_all=True, append_images=sayfalar[1:])
    assert len(PdfReader(gorsel.donustur(str(tmp_path / "t.tiff"), "pdf", str(tmp_path))).pages) == 3


def test_buyuk_foto_pdf_sayfasi_makul(tmp_path):
    from pypdf import PdfReader

    Image.new("RGB", (4000, 3000), "white").save(tmp_path / "b.jpg")
    sayfa = PdfReader(gorsel.donustur(str(tmp_path / "b.jpg"), "pdf", str(tmp_path))).pages[0]
    genislik_inc = float(sayfa.mediabox.width) / 72
    assert genislik_inc <= 8.3 + 0.01  # A4 genişliğini aşmaz (eskiden ~42 inç)


def test_kucuk_ico(tmp_path):
    Image.new("RGBA", (8, 8), "red").save(tmp_path / "k.png")
    cikti = Path(gorsel.donustur(str(tmp_path / "k.png"), "ico", str(tmp_path)))
    assert cikti.stat().st_size > 100
    assert Image.open(cikti).size == (16, 16)


def test_renk_profili_korunur(tmp_path):
    from PIL import ImageCms

    profil = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
    Image.new("RGB", (10, 10), "red").save(tmp_path / "p.jpg", icc_profile=profil)
    for hedef in ("png", "webp"):
        cikti = Image.open(gorsel.donustur(str(tmp_path / "p.jpg"), hedef, str(tmp_path / "o")))
        assert cikti.info.get("icc_profile") == profil


def test_saydamlik_beyaz_arka_plan(tmp_path):
    Image.new("RGBA", (10, 10), (0, 0, 0, 0)).save(tmp_path / "s.png")
    cikti = Image.open(gorsel.donustur(str(tmp_path / "s.png"), "jpg", str(tmp_path)))
    assert min(cikti.getpixel((5, 5))) > 245


def test_bozuk_gorsel_yarim_cikti_birakmaz(tmp_path):
    (tmp_path / "bozuk.png").write_bytes(b"\x89PNG\r\n\x1a\n" + b"\0" * 50)
    with pytest.raises((OSError, SyntaxError)):
        gorsel.donustur(str(tmp_path / "bozuk.png"), "jpg", str(tmp_path / "o"))
    assert not (tmp_path / "o").exists() or not any((tmp_path / "o").iterdir())


def test_animasyonlu_kareler_iterasyonu_sabit(tmp_path):
    palet = [0, 0, 0, 255, 0, 0, 0, 255, 0, 0, 0, 255] + [0] * 756
    kareler = []
    for i in range(4):
        kare = Image.new("P", (10, 10), i)
        kare.putpalette(palet)
        kareler.append(kare)
    kareler[0].save(tmp_path / "p.gif", save_all=True, append_images=kareler[1:])
    assert Image.open(tmp_path / "p.gif").n_frames == 4
    webp = Image.open(gorsel.donustur(str(tmp_path / "p.gif"), "webp", str(tmp_path)))
    assert len(list(ImageSequence.Iterator(webp))) == 4
