"""Güvenlik gerileme testleri: kapatılan açıkların tekrar açılmadığını denetler."""

import base64
from pathlib import Path

import pytest
from conftest import ffmpeg, png_bayt

from converters import belge, gorsel, ortak, ses
from converters import belge_bloklar as bl


def test_data_uri_adi_ile_yol_gecisi_yok(tmp_path, monkeypatch):
    """Eskiden HTML -> MD, 'data:image/..\\..\\x.bat' adıyla istenen klasöre
    dosya yazabiliyordu (ör. Başlangıç klasörüne -> kod çalıştırma)."""
    monkeypatch.setattr(belge, "_word_kullanilabilir", lambda: False)
    ic = tmp_path / "a" / "b"
    ic.mkdir(parents=True)
    veri = base64.b64encode(png_bayt()).decode()
    (ic / "kotu.html").write_text(
        "<p>x</p>"
        f'<img src="data:image/..\\..\\..\\KACAK.bat;base64,{veri}">'
        f'<img src="data:image/../../KACAK2.bat;base64,{veri}">', encoding="utf-8")
    belge.donustur(str(ic / "kotu.html"), "md")
    yazilanlar = [p for p in tmp_path.rglob("*") if p.is_file()]
    assert not [p for p in yazilanlar if "KACAK" in p.name]
    for p in yazilanlar:
        if p.suffix == ".png":
            assert p.parent.name == "kotu_dosyalar"


def test_resim_adi_temizleme():
    assert bl.guvenli_dosya_adi("..\\..\\x.bat") == "x.bat.bin"
    assert bl.guvenli_dosya_adi("../../evil.png") == "evil.png"
    assert bl.guvenli_dosya_adi("C:\\Windows\\a.jpg") == "a.jpg"
    assert bl.guvenli_dosya_adi("CON.png") == "_CON.png"
    assert bl.guvenli_dosya_adi("") == "resim.png"
    assert bl.guvenli_dosya_adi('a<>:"|?*.png') == "a_______.png"


def test_md_html_yerel_dosya_sizdirmaz(tmp_path, monkeypatch):
    """Eskiden '![x](C:/Windows/win.ini)' dosyanın içeriğini çıktıya gömüyordu."""
    monkeypatch.setattr(belge, "_word_kullanilabilir", lambda: False)
    gizli = tmp_path / "gizli.txt"
    gizli.write_text("PAROLA=123", encoding="utf-8")
    (tmp_path / "gercek.png").write_bytes(png_bayt())
    (tmp_path / "s.md").write_text(
        f"![a]({gizli.as_posix()})\n\n![b](C:/Windows/win.ini)\n\n![c](gercek.png)\n",
        encoding="utf-8")
    cikti = Path(belge.donustur(str(tmp_path / "s.md"), "html", str(tmp_path / "o")))
    icerik = cikti.read_text(encoding="utf-8")
    assert base64.b64encode(b"PAROLA=123").decode() not in icerik
    assert icerik.count("<img") == 1  # yalnızca gerçek resim gömülür


@pytest.mark.parametrize("adres", [
    r"\\192.0.2.1\pay\x.png", "//192.0.2.1/pay/x.png", r"/\192.0.2.1\pay\x.png",
    "file://192.0.2.1/pay/x.png", "file:////192.0.2.1/pay/x.png",
    r"\\?\UNC\192.0.2.1\pay\x.png", "http://192.0.2.1/x.png", "https://ornek.com/x.png",
    "javascript:alert(1)",
])
def test_ag_ve_uzak_resim_adresleri_reddedilir(tmp_path, monkeypatch, adres):
    """UNC yolları dosya sistemine hiç sorulmadan reddedilmeli (sorulsa bile
    Windows NTLM kimlik özetini o sunucuya gönderir)."""
    sorgulanan = []
    gercek = Path.resolve

    def izle(self, *a, **k):
        sorgulanan.append(str(self))
        return gercek(self, *a, **k)

    monkeypatch.setattr(Path, "resolve", izle)
    assert bl.yerel_resim_yolu(adres, tmp_path) is None
    assert not [s for s in sorgulanan if "192.0.2.1" in s]


def test_yerel_resim_kabul_edilir(tmp_path):
    (tmp_path / "klasör adı").mkdir()
    (tmp_path / "klasör adı" / "r ş.png").write_bytes(png_bayt())
    assert bl.yerel_resim_yolu("klasör adı/r ş.png", tmp_path)
    assert bl.yerel_resim_yolu("klas%C3%B6r%20ad%C4%B1/r%20%C5%9F.png", tmp_path)
    assert bl.yerel_resim_yolu((tmp_path / "klasör adı" / "r ş.png").as_uri(), tmp_path)
    (tmp_path / "metin.png").write_text("resim değil", encoding="utf-8")
    assert bl.yerel_resim_yolu("metin.png", tmp_path) is None
    assert bl.yerel_resim_yolu(".", tmp_path) is None  # klasör


def test_ag_adresli_html_worde_verilmez(tmp_path, monkeypatch):
    assert bl.ag_adresi_var('<img src="\\\\sunucu\\pay\\a.png">')
    assert bl.ag_adresi_var("<link href='//sunucu/a.css'>")
    assert bl.ag_adresi_var('<div style="background:url(\\\\sunucu\\a.png)">')
    assert bl.ag_adresi_var('<style>@import "//sunucu/a.css";</style>')
    assert bl.ag_adresi_var('<img src="file://sunucu/pay/a.png">')
    assert not bl.ag_adresi_var('<img src="resim.png"><a href="https://ornek.com">x</a>')
    assert not bl.ag_adresi_var('<img src="file:///C:/a.png">')

    cagrildi = []
    monkeypatch.setattr(belge, "_word_kullanilabilir", lambda: True)
    monkeypatch.setattr(belge, "_word_ile_donustur", lambda *a, **k: cagrildi.append(a))
    (tmp_path / "t.html").write_text(
        '<h1>x</h1><img src="\\\\192.0.2.1\\pay\\a.png">', encoding="utf-8")
    belge.donustur(str(tmp_path / "t.html"), "docx", str(tmp_path / "o"))
    assert not cagrildi


def test_toplu_iste_kaynak_ezilmez(tmp_path):
    """Eskiden a.png + a.jpg -> JPG yapınca a.png'nin çıktısı kaynak a.jpg'yi ezerdi."""
    from PIL import Image

    Image.new("RGB", (10, 10), "blue").save(tmp_path / "a.png")
    Image.new("RGB", (10, 10), "green").save(tmp_path / "a.jpg")
    once = (tmp_path / "a.jpg").read_bytes()
    c1 = gorsel.donustur(str(tmp_path / "a.png"), "jpg")
    c2 = gorsel.donustur(str(tmp_path / "a.jpg"), "jpg")
    assert (tmp_path / "a.jpg").read_bytes() == once
    assert len({c1, c2, str(tmp_path / "a.jpg")}) == 3


def test_var_olan_dosya_ezilmez(tmp_path):
    (tmp_path / "rapor.txt").write_text("kaynak", encoding="utf-8")
    (tmp_path / "rapor.md").write_text("KULLANICININ DOSYASI", encoding="utf-8")
    cikti = belge.donustur(str(tmp_path / "rapor.txt"), "md")
    assert (tmp_path / "rapor.md").read_text(encoding="utf-8") == "KULLANICININ DOSYASI"
    assert Path(cikti).name == "rapor_donusturuldu.md"


def test_ffmpeg_ag_protokolu_kapali(tmp_path):
    """Medya dosyası kılığındaki bir HLS listesi ffmpeg'e ağdan veri çektiremez."""
    (tmp_path / "tuzak.mp3").write_text(
        "#EXTM3U\n#EXT-X-TARGETDURATION:1\n#EXTINF:1,\nhttp://127.0.0.1:9/a.ts\n"
        "#EXT-X-ENDLIST\n", encoding="utf-8")
    with pytest.raises(RuntimeError) as hata:
        ses.donustur(str(tmp_path / "tuzak.mp3"), "wav", str(tmp_path / "o"))
    assert "whitelist" in str(hata.value) or "akış" in str(hata.value) \
        or "okunamadı" in str(hata.value)
    assert not list((tmp_path / "o").glob("*")), "yarım çıktı silinmeli"


def test_ffmpeg_dosya_adi_protokol_sayilmaz(tmp_path):
    ffmpeg("-f", "lavfi", "-i", "sine=duration=0.3", tmp_path / "concat_a.wav")
    cikti = ses.donustur(str(tmp_path / "concat_a.wav"), "mp3", str(tmp_path))
    assert Path(cikti).stat().st_size > 0


def test_hata_metinleri_turkce():
    import errno

    assert "bulunamadı" in ortak.hata_metni(FileNotFoundError(errno.ENOENT, "x"))
    assert "erişilemedi" in ortak.hata_metni(PermissionError(errno.EACCES, "x"))
    assert "yer kalmadı" in ortak.hata_metni(OSError(errno.ENOSPC, "x"))
    from PIL import UnidentifiedImageError

    assert "tanınamadı" in ortak.hata_metni(UnidentifiedImageError("x"))
