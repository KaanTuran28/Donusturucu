"""Belge dönüşümleri: tüm format çiftleri (Word ile ve Word olmadan) +
tek tek düzeltilmiş hataların gerileme testleri."""

import zipfile
from pathlib import Path

import pytest
from conftest import (ISARET, WORD_VAR, metni_oku, png_bayt, resim_var_mi,
                      word_gerekli)

from converters import belge
from converters import belge_bloklar as bl

KAYNAKLAR = ["pdf", "docx", "txt", "html", "md", "rtf", "odt"]
HEDEFLER = belge.FORMATLAR


@pytest.fixture(params=["dahili", "word"])
def motor(request, monkeypatch):
    """Her testi hem dahili motorla hem (kuruluysa) Word ile çalıştırır."""
    if request.param == "word":
        if not WORD_VAR:
            pytest.skip("Microsoft Word kurulu değil")
    else:
        monkeypatch.setattr(belge, "_word_kullanilabilir", lambda: False)
    return request.param


@pytest.mark.parametrize("hedef", HEDEFLER)
@pytest.mark.parametrize("kaynak", KAYNAKLAR)
def test_tum_format_ciftleri(belge_ornekleri, tmp_path, motor, kaynak, hedef):
    cikti = Path(belge.donustur(str(belge_ornekleri[kaynak]), hedef, str(tmp_path)))
    assert cikti.exists() and cikti.stat().st_size > 0
    assert cikti.suffix == "." + hedef
    assert ISARET in metni_oku(cikti), "Türkçe metin kaybolmuş/bozulmuş"


@word_gerekli
@pytest.mark.parametrize("hedef", ["pdf", "docx", "txt", "md", "odt"])
def test_eski_doc_kaynak(belge_ornekleri, tmp_path, hedef):
    cikti = Path(belge.donustur(str(belge_ornekleri["doc"]), hedef, str(tmp_path)))
    assert ISARET in metni_oku(cikti)


def test_doc_word_yoksa_anlasilir_hata(belge_ornekleri, tmp_path, monkeypatch):
    doc = tmp_path / "eski.doc"
    doc.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 504)
    monkeypatch.setattr(belge, "_word_kullanilabilir", lambda: False)
    with pytest.raises(RuntimeError, match="Microsoft Word"):
        belge.donustur(str(doc), "pdf", str(tmp_path))


@pytest.mark.parametrize("hedef", ["docx", "odt", "pdf", "html", "md"])
@pytest.mark.parametrize("kaynak", ["docx", "html", "md", "odt"])
def test_resimler_korunur(belge_ornekleri, tmp_path, motor, kaynak, hedef):
    if kaynak == hedef:
        pytest.skip("aynı format kopyalanır")
    cikti = Path(belge.donustur(str(belge_ornekleri[kaynak]), hedef, str(tmp_path)))
    assert resim_var_mi(cikti)
    if hedef == "docx":  # resim dış bağlantı değil, gerçekten gömülü olmalı
        rels = zipfile.ZipFile(cikti).read("word/_rels/document.xml.rels").decode()
        assert 'TargetMode="External"' not in rels


def test_ayni_format_kayipsiz_kopya(belge_ornekleri, tmp_path):
    for f in KAYNAKLAR:
        cikti = Path(belge.donustur(str(belge_ornekleri[f]), f, str(tmp_path)))
        assert cikti.read_bytes() == belge_ornekleri[f].read_bytes()


def test_ayni_klasore_ayni_format_kaynagi_ezmez(tmp_path):
    kaynak = tmp_path / "a.txt"
    kaynak.write_text("orijinal", encoding="utf-8")
    cikti1 = Path(belge.donustur(str(kaynak), "txt"))
    cikti2 = Path(belge.donustur(str(kaynak), "txt"))
    assert kaynak.read_text(encoding="utf-8") == "orijinal"
    assert cikti1.name == "a_donusturuldu.txt"
    assert cikti2.name == "a_donusturuldu_2.txt"


# ------------------------------------------------------------ okuyucu hataları

def test_docx_kopru_ve_icerik_denetimi_metni(belge_ornekleri, tmp_path, monkeypatch):
    monkeypatch.setattr(belge, "_word_kullanilabilir", lambda: False)
    metin = metni_oku(Path(belge.donustur(str(belge_ornekleri["docx"]), "md", str(tmp_path))))
    assert "KÖPRÜMETNİ" in metin  # köprü metni eskiden siliniyordu
    assert "h00 | h01" in metin   # tablo


def test_docx_sdt_icerigi_okunur(tmp_path):
    import docx
    from docx.oxml import parse_xml

    d = docx.Document()
    d.add_paragraph("önce")
    d.element.body.insert(1, parse_xml(
        '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:sdtContent><w:p><w:r><w:t>KAPAKSAYFASI</w:t></w:r></w:p></w:sdtContent></w:sdt>"))
    d.save(tmp_path / "s.docx")
    assert "KAPAKSAYFASI" in [b.metin() for b in bl.oku_docx(tmp_path / "s.docx")]


def test_odt_bolum_tablo_kalin(belge_ornekleri):
    bloklar = bl.oku_odt(belge_ornekleri["odt"])
    metinler = [b.metin() for b in bloklar]
    assert ISARET in metinler              # text:section içindeki paragraf
    assert "h00 | h01" in metinler         # tablo
    assert any(p.kalin and p.metin == "kalın" for b in bloklar for p in b.parcalar)
    assert [b.tip for b in bloklar if b.metin().startswith("madde")] == ["li", "li"]


def test_odt_media_klasorundeki_resimler(belge_ornekleri, tmp_path):
    """Word ODT'de resimleri "media/" altına koyar; odfpy bunları yüklemez."""
    kaynak = zipfile.ZipFile(belge_ornekleri["odt"])
    resim_adi = next(n for n in kaynak.namelist() if n.startswith("Pictures/"))
    with zipfile.ZipFile(tmp_path / "w.odt", "w") as hedef:
        for ad in kaynak.namelist():
            veri = kaynak.read(ad)
            if ad in ("content.xml", "META-INF/manifest.xml"):
                veri = veri.replace(resim_adi.encode(), b"media/image1.png")
            hedef.writestr("media/image1.png" if ad == resim_adi else ad, veri)
    assert any(b.tip == "resim" for b in bl.oku_odt(tmp_path / "w.odt"))


def test_ansi_txt_turkce(tmp_path):
    (tmp_path / "ansi.txt").write_bytes(ISARET.encode("cp1254"))
    assert bl.oku_txt(tmp_path / "ansi.txt")[0].metin() == ISARET


def test_html_meta_charset(tmp_path):
    (tmp_path / "w.html").write_bytes(
        f'<html><head><meta http-equiv="Content-Type" content="text/html; '
        f'charset=windows-1254"></head><body><p>{ISARET}</p></body></html>'.encode("cp1254"))
    assert bl.oku_html(tmp_path / "w.html")[0].metin() == ISARET


def test_html_liste_tablo_pre():
    bloklar = bl.oku_html_metin(
        "<ul><li><p>bir</p></li><li><p>iki</p></li></ul>"
        "<table><tr><td>a</td><td>b</td></tr></table>"
        "<pre>x = 1\n    y = 2</pre>")
    ciftler = [(b.tip, b.metin()) for b in bloklar if b.metin().strip()]
    assert ciftler[:2] == [("li", "bir"), ("li", "iki")]  # <li><p> madde kalır
    assert ("p", "a | b") in ciftler                        # hücreler ayrılır
    assert ("p", "    y = 2") in ciftler                     # kod girintisi korunur
    tipler = [b.tip for b in bloklar]
    assert tipler[0:2] == ["li", "li"], "maddeler arasına boş satır girmemeli"


def test_pdf_kullanici_parolasi_anlasilir_hata(tmp_path, monkeypatch):
    from pypdf import PdfWriter

    yazici = PdfWriter()
    yazici.add_blank_page(200, 200)
    yazici.encrypt(user_password="gizli", owner_password="sahip")
    with open(tmp_path / "p.pdf", "wb") as f:
        yazici.write(f)
    monkeypatch.setattr(belge, "_word_kullanilabilir", lambda: False)
    with pytest.raises(RuntimeError, match="parola"):
        belge.donustur(str(tmp_path / "p.pdf"), "txt", str(tmp_path / "o"))
    assert not list((tmp_path / "o").glob("*")), "yarım çıktı silinmeli"


def test_pdf_yalniz_sahip_parolasi_acilir(belge_ornekleri, tmp_path):
    from pypdf import PdfReader, PdfWriter

    yazici = PdfWriter(clone_from=PdfReader(str(belge_ornekleri["pdf"])))
    yazici.encrypt(user_password="", owner_password="sahip")
    with open(tmp_path / "s.pdf", "wb") as f:
        yazici.write(f)
    assert ISARET in metni_oku(Path(belge.donustur(str(tmp_path / "s.pdf"), "txt", str(tmp_path))))


@word_gerekli
def test_parolali_docx_takilmaz(tmp_path):
    """Eskiden Word görünmez bir parola kutusunda sonsuza kadar bekliyordu."""
    import time

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    try:
        word.DisplayAlerts = 0
        d = word.Documents.Add()
        d.Content.Text = "gizli"
        d.Password = "abc"
        d.SaveAs2(str(tmp_path / "p.docx"), FileFormat=16)
        d.Close(0)
    finally:
        word.Quit()
    assert belge.parolali_mi(tmp_path / "p.docx")
    baslangic = time.monotonic()
    with pytest.raises(RuntimeError, match="parola"):
        belge.donustur(str(tmp_path / "p.docx"), "pdf", str(tmp_path / "o"))
    assert time.monotonic() - baslangic < 5


def test_parolali_odt_tespiti(tmp_path):
    with zipfile.ZipFile(tmp_path / "p.odt", "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.text")
        z.writestr("META-INF/manifest.xml",
                   '<manifest:manifest><manifest:file-entry><manifest:encryption-data/>'
                   "</manifest:file-entry></manifest:manifest>")
    assert belge.parolali_mi(tmp_path / "p.odt")
    with pytest.raises(RuntimeError, match="parola"):
        belge.donustur(str(tmp_path / "p.odt"), "txt", str(tmp_path))


# ------------------------------------------------------------ yazıcı hataları

def test_kontrol_karakterleri_docx_odt_cokertmez(tmp_path):
    (tmp_path / "k.txt").write_text("a\x00b\x01c\x0cd\x1fe", encoding="utf-8")
    for hedef in ("docx", "odt", "pdf", "rtf", "html"):
        cikti = Path(belge.donustur(str(tmp_path / "k.txt"), hedef, str(tmp_path)))
        assert "ab" in metni_oku(cikti).replace(" ", "")
    # ODT geçerli XML olmalı
    from odf.opendocument import load

    load(str(tmp_path / "k.odt"))


def test_uzun_resim_pdf_cokertmez(tmp_path):
    bloklar = [bl.Blok("p", [bl.Parca("x")]),
               bl.Blok("resim", resim=png_bayt((800, 5000)), resim_ad="uzun.png")]
    bl.yaz_pdf(bloklar, tmp_path / "u.pdf")
    assert resim_var_mi(tmp_path / "u.pdf")


def test_desteklenmeyen_resim_turleri_donusturulur(tmp_path):
    import io

    from PIL import Image

    webp, cmyk = io.BytesIO(), io.BytesIO()
    Image.new("RGB", (40, 30), "blue").save(webp, "WEBP")
    Image.new("CMYK", (40, 30), (0, 100, 100, 0)).save(cmyk, "JPEG")
    bloklar = [bl.Blok("resim", resim=webp.getvalue(), resim_ad="a.webp"),
               bl.Blok("resim", resim=cmyk.getvalue(), resim_ad="b.jpg")]
    bl.yaz_docx(bloklar, tmp_path / "r.docx")
    bl.yaz_pdf(bloklar, tmp_path / "r.pdf")
    assert len([n for n in zipfile.ZipFile(tmp_path / "r.docx").namelist()
                if n.startswith("word/media/")]) == 2
    assert resim_var_mi(tmp_path / "r.pdf")


def test_rtf_unicode_ve_emoji(tmp_path):
    bl.yaz_rtf([bl.Blok("p", [bl.Parca("a😀bＡş")])], tmp_path / "e.rtf")
    icerik = (tmp_path / "e.rtf").read_text(encoding="ascii")
    assert "\\u-10179?\\u-8704?" in icerik  # vekil çift, işaretli 16-bit
    assert "\\u-223?" in icerik              # U+FF21 > 32767 -> negatif
    assert "\\u351?" in icerik               # ş
    from striprtf.striprtf import rtf_to_text

    assert "ş" in rtf_to_text(icerik)


def test_md_yazici_ozel_karakterleri_kacirir(tmp_path):
    bloklar = [bl.Blok("p", [bl.Parca("# başlık değil")]),
               bl.Blok("p", [bl.Parca("1. liste değil")]),
               bl.Blok("p", [bl.Parca("<script>alert(1)</script> ve *yıldız* [x](y)")]),
               bl.Blok("p", [bl.Parca("snake_case_ad")])]
    bl.yaz_md(bloklar, tmp_path / "k.md")
    html = bl.md_html((tmp_path / "k.md").read_text(encoding="utf-8"))
    assert "<h1>" not in html and "<ol>" not in html
    assert "<script>" not in html and "<em>" not in html and "<a " not in html
    assert "snake_case_ad" in html


def test_md_ayni_adli_resimler_ezilmez(tmp_path):
    bloklar = [bl.Blok("resim", resim=png_bayt(renk=r), resim_ad="resim.png")
               for r in ("red", "green", "blue")]
    bl.yaz_md(bloklar, tmp_path / "Rapor Dosyası.md")
    klasor = tmp_path / "Rapor Dosyası_dosyalar"
    assert len(list(klasor.iterdir())) == 3
    # boşluklu klasör adıyla bağlantılar çalışmalı (MD -> HTML -> resimler okunur)
    geri = bl.oku_md(tmp_path / "Rapor Dosyası.md")
    assert sum(1 for b in geri if b.tip == "resim") == 3
