"""Görsel dönüştürücü - Pillow tabanlı.

Desteklenen formatlar: PNG, JPG, BMP, GIF, WEBP, TIFF, ICO, AVIF, HEIC
Ayrıca görselden PDF üretimi desteklenir.
"""

import io
import shutil
from pathlib import Path

from PIL import Image, ImageOps, ImageSequence

from . import ortak

try:  # HEIC/HEIF okuma-yazma desteği
    from pillow_heif import register_heif_opener

    register_heif_opener()
    _HEIC_VAR = True
except ImportError:
    _HEIC_VAR = False

FORMATLAR = ["png", "jpg", "jpeg", "bmp", "gif", "webp", "tiff", "ico", "avif", "pdf"]
if _HEIC_VAR:
    FORMATLAR.insert(-1, "heic")

# Pillow'un beklediği format adları
_PIL_FORMAT = {
    "png": "PNG",
    "jpg": "JPEG",
    "jpeg": "JPEG",
    "bmp": "BMP",
    "gif": "GIF",
    "webp": "WEBP",
    "tiff": "TIFF",
    "ico": "ICO",
    "avif": "AVIF",
    "heic": "HEIF",
    "pdf": "PDF",
}

# Saydamlık (alfa kanalı) desteklemeyen formatlar
_ALFA_YOK = {"jpg", "jpeg", "bmp", "pdf"}

# Tüm kareleri korunan eşleşmeler: animasyonlu GIF/WEBP/PNG -> GIF/WEBP/PNG
# (APNG) animasyonlu kalır; çok sayfalı TIFF -> PDF/TIFF tüm sayfaları içerir.
# Diğer durumlarda (ör. GIF -> PDF, TIFF -> JPG) yalnızca ilk kare kullanılır.
_ANIMASYON_KAYNAK = {"GIF", "WEBP", "PNG", "AVIF"}
_ANIMASYON_HEDEF = {"gif", "webp", "png"}
_SAYFALI_KAYNAK = {"TIFF"}
_SAYFALI_HEDEF = {"tiff", "pdf"}

# Her hedefin olduğu gibi yazabildiği renk kipleri; diğerleri RGB/RGBA'ya
# çevrilir. (Eskiden CMYK bir JPG -> PNG "cannot write mode CMYK as PNG"
# hatasıyla çöküyordu.)
_KIPLER = {
    "png": {"1", "L", "LA", "I;16", "RGB", "RGBA"},
    "gif": {"L", "RGB", "RGBA"},
    "tiff": {"1", "L", "LA", "I;16", "RGB", "RGBA", "CMYK"},
}


def _renk_ailesi(kip: str) -> str:
    if kip == "CMYK":
        return "CMYK"
    if kip in ("1", "L", "LA", "I", "F") or kip.startswith("I;16"):
        return "L"
    return "RGB"


def _sekiz_bite_indir(kare: Image.Image) -> Image.Image:
    """16/32-bit tamsayı ve kayan noktalı gri görselleri 8-bit'e ölçekleyerek
    indirir. Pillow'un doğrudan convert("RGB")'si 255 üstünü kırptığı için
    16-bit görseller bembeyaz çıkıyordu."""
    if kare.mode == "I" or kare.mode.startswith("I;16"):
        tam = kare if kare.mode == "I" else kare.convert("I")
        _, en_buyuk = tam.getextrema()
        if en_buyuk > 255:
            olcek = 255 / (65535 if en_buyuk <= 65535 else en_buyuk)
            tam = tam.point(lambda v: v * olcek)
        return tam.convert("L")
    if kare.mode == "F":
        _, en_buyuk = kare.getextrema()
        olcek = 255.0 if en_buyuk <= 1.0 else 1.0
        return kare.point(lambda v: v * olcek).convert("L")
    return kare


def _cmyk_rgb(kare: Image.Image) -> Image.Image:
    """CMYK'yı, varsa gömülü renk profiliyle doğru renklerle sRGB'ye çevirir."""
    profil = kare.info.get("icc_profile")
    if profil:
        try:
            from PIL import ImageCms

            return ImageCms.profileToProfile(
                kare, ImageCms.ImageCmsProfile(io.BytesIO(profil)),
                ImageCms.createProfile("sRGB"), outputMode="RGB")
        except Exception:
            pass  # profil okunamazsa basit dönüşüme düş
    return kare.convert("RGB")


def _kip_duzelt(kare: Image.Image, hedef: str) -> Image.Image:
    """Kareyi hedef formatın yazabileceği bir renk kipine getirir."""
    if not (kare.mode.startswith("I;16") and "I;16" in _KIPLER.get(hedef, ())):
        kare = _sekiz_bite_indir(kare)
    if kare.mode == "CMYK" and "CMYK" not in _KIPLER.get(hedef, ()):
        kare = _cmyk_rgb(kare)

    if hedef in _ALFA_YOK:
        if kare.has_transparency_data:
            # Saydam alanlara beyaz arka plan uygula
            arka = Image.new("RGB", kare.size, (255, 255, 255))
            rgba = kare.convert("RGBA")
            arka.paste(rgba, mask=rgba.getchannel("A"))
            return arka
        return kare if kare.mode == "RGB" else kare.convert("RGB")

    if kare.mode in ("P", "PA"):
        kare = kare.convert("RGBA")
    if kare.mode not in _KIPLER.get(hedef, {"RGB", "RGBA"}):
        kare = kare.convert("RGBA" if kare.has_transparency_data else "RGB")
    return kare


def _pdf_cozunurlugu(boyut: tuple[int, int]) -> float:
    """Görseli PDF'te en fazla A4 boyutunda bir sayfaya sığdıran DPI.
    (Sabit 96 DPI ile 4000px'lik bir fotoğraf 1 metrelik sayfa oluyordu.)"""
    genislik, yukseklik = boyut
    return max(96.0, genislik / 8.27, yukseklik / 11.69)


def donustur(kaynak: str, hedef_format: str, cikti_klasoru: str | None = None,
             iptal=None) -> str:
    """Bir görsel dosyasını hedef formata dönüştürür, çıktı yolunu döndürür."""
    kaynak_yol = Path(kaynak)
    hedef_format = hedef_format.lower().lstrip(".")
    if hedef_format not in _PIL_FORMAT:
        raise ValueError(f"Desteklenmeyen görsel formatı: {hedef_format}")
    if hedef_format == "heic" and not _HEIC_VAR:
        raise RuntimeError("HEIC desteği için 'pillow-heif' paketi gerekli.")
    ortak.kaynak_denetle(kaynak_yol)

    with Image.open(kaynak_yol) as im, \
            ortak.cikti_dosyasi(kaynak_yol, hedef_format, cikti_klasoru) as cikti:
        if im.format == _PIL_FORMAT[hedef_format]:
            # İçerik zaten hedef formatta: motordan geçirmeden kayıpsız
            # kopyala (jpg/jpeg eşdeğerliğini de kapsar; animasyonlu GIF
            # gibi çok kareli görsellerde kare kaybını da önler).
            shutil.copy2(kaynak_yol, cikti)
            return str(cikti)
        _kaydet(im, cikti, hedef_format)
    return str(cikti)


def _kaydet(im: Image.Image, cikti: Path, hedef_format: str):
    kaynak_kip = im.mode
    icc = im.info.get("icc_profile")
    kare_sayisi = getattr(im, "n_frames", 1)

    kaydet_args: dict = {}
    if hedef_format in ("jpg", "jpeg"):
        kaydet_args["quality"] = 95
    elif hedef_format in ("avif", "heic", "webp"):
        kaydet_args["quality"] = 90

    cok_kare = kare_sayisi > 1 and (
        (im.format in _ANIMASYON_KAYNAK and hedef_format in _ANIMASYON_HEDEF)
        or (im.format in _SAYFALI_KAYNAK and hedef_format in _SAYFALI_HEDEF))
    if cok_kare:
        # Tüm kareleri/sayfaları koru
        kareler, sureler = [], []
        for kare in ImageSequence.Iterator(im):
            sureler.append(kare.info.get("duration", 100))
            kareler.append(_kip_duzelt(kare.copy(), hedef_format))
        ilk = kareler[0]
        kaydet_args.update(save_all=True, append_images=kareler[1:])
        if hedef_format in _ANIMASYON_HEDEF:
            kaydet_args.update(duration=sureler, loop=im.info.get("loop", 0))
        if hedef_format == "pdf":
            kaydet_args["resolution"] = _pdf_cozunurlugu(ilk.size)
    else:
        # Tek kare: telefon fotoğraflarındaki EXIF yön bilgisini uygula
        # (yoksa dikey çekilmiş fotoğraflar yan yatmış çıkıyordu).
        ilk = _kip_duzelt(ImageOps.exif_transpose(im), hedef_format)
        if hedef_format == "pdf":
            kaydet_args["resolution"] = _pdf_cozunurlugu(ilk.size)
        elif hedef_format == "ico":
            # Pillow'un ICO yazıcısı 16px'in altındaki görsellerde sessizce
            # boş/geçersiz (6 bayt) bir dosya üretir -- bu yüzden asgari
            # 16px'e tamamlanır. ICO en fazla 256x256 destekler. Standart
            # çoklu-boyut listesi de açıkça verilir (Pillow'un örtük
            # varsayılanına güvenmek yerine).
            en_buyuk = max(ilk.size)
            if en_buyuk < 16:
                oran = 16 / en_buyuk
                ilk = ilk.resize((max(16, round(ilk.width * oran)),
                                  max(16, round(ilk.height * oran))))
            elif en_buyuk > 256:
                ilk.thumbnail((256, 256))
            kaydet_args["sizes"] = [
                (s, s) for s in (16, 24, 32, 48, 64, 128, 256) if s <= max(ilk.size)]

    # Renk profilini koru (ör. iPhone'un Display P3 fotoğrafları profilsiz
    # kaydedilince soluk görünür) -- ama yalnızca renk uzayı değişmediyse:
    # CMYK/gri profili RGB veriye eklemek renkleri bozar.
    if icc and _renk_ailesi(kaynak_kip) == _renk_ailesi(ilk.mode):
        kaydet_args["icc_profile"] = icc

    ilk.save(cikti, _PIL_FORMAT[hedef_format], **kaydet_args)
