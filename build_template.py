# -*- coding: utf-8 -*-
"""Kök Neden Analizi / RW raporu PowerPoint şablonu üretir."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

OUT = "KOK_NEDEN_ANALIZI_SABLONU.pptx"

# Renkler
DARK = RGBColor(0x1F, 0x2A, 0x33)    # koyu antrasit
SLATE = RGBColor(0x2E, 0x4A, 0x5C)   # çelik mavisi
LIGHT = RGBColor(0xEE, 0xF1, 0xF4)   # açık gri kart
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE = RGBColor(0xE8, 0x77, 0x2E)  # kök neden / vurgu
GREEN = RGBColor(0x2E, 0x7D, 0x6B)
RED = RGBColor(0xC0, 0x39, 0x2B)
AMBER = RGBColor(0xD4, 0xA0, 0x17)
MUTED = RGBColor(0x5F, 0x6E, 0x7C)
LINE = RGBColor(0xC9, 0xD1, 0xD8)
HEAD = "Arial"
BODY = "Calibri"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
TOTAL_STEPS = 10


# ---------- yardımcılar ----------
def box(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None, dash=False, lw=1):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
        if dash:
            s.line.dash_style = 4  # dash
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    s.shadow.inherit = False
    s.text_frame.text = ""
    return s


def text(slide, x, y, w, h, runs, size=14, color=DARK, bold=False, italic=False, font=BODY,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, margin=0.0, shape=None, space_after=0):
    """runs: str | list of paragraphs; paragraph = str | list of (text, {opts})"""
    tb = shape if shape is not None else slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, Inches(margin))
    paras = runs if isinstance(runs, list) else [runs]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        if space_after:
            para.space_after = Pt(space_after)
        segs = p if isinstance(p, list) else [(p, {})]
        for seg, o in segs:
            r = para.add_run()
            r.text = seg
            f = r.font
            f.name = o.get("font", font)
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", italic)
            f.color.rgb = o.get("color", color)
    return tb


def bullets(slide, x, y, w, h, items, size=14, color=DARK):
    tb = text(slide, x, y, w, h, items, size=size, color=color, space_after=4)
    for p in tb.text_frame.paragraphs:
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(Inches(0.22)))
        pPr.set("indent", str(-Inches(0.22)))
        bu = etree.SubElement(pPr, qn("a:buChar"))
        bu.set("char", "•")
    return tb


def badge(slide, n):
    c = box(slide, 0.6, 0.5, 0.65, 0.65, fill=ORANGE, shape=MSO_SHAPE.OVAL)
    text(slide, 0, 0, 0, 0, f"{n:02d}", size=16, bold=True, color=WHITE, font=HEAD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=c)


def content_slide(n, title, hint, notes):
    s = prs.slides.add_slide(BLANK)
    badge(s, n)
    text(s, 1.45, 0.45, 11.3, 0.75, title, size=30, bold=True, font=HEAD, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 1.45, 1.17, 11.3, 0.4, hint, size=12, italic=True, color=MUTED)
    # alt bilgi
    text(s, 0.6, 6.98, 8.0, 0.3, [[("KÖK NEDEN ANALİZİ RAPORU", {"bold": True}),
                                   ("   ·   [Rapor No]   ·   [Konu Başlığı]", {})]],
         size=10, color=MUTED)
    text(s, 11.73, 6.98, 1.0, 0.3, f"{n + 1} / {TOTAL_STEPS + 1}", size=10, color=MUTED, align=PP_ALIGN.RIGHT)
    s.notes_slide.notes_text_frame.text = notes
    return s


def icon_circle(slide, x, y, d, glyph, fill):
    c = box(slide, x, y, d, d, fill=fill, shape=MSO_SHAPE.OVAL)
    text(slide, 0, 0, 0, 0, glyph, size=int(d * 30), bold=True, color=WHITE, font=HEAD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=c)


def photo_placeholder(slide, x, y, w, h, label):
    b = box(slide, x, y, w, h, fill=LIGHT, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04, dash=True, lw=1.5)
    text(slide, 0, 0, 0, 0, [[("▣", {"size": 28, "color": LINE})], [(label, {})]], size=13, italic=True,
         color=MUTED, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=b)


def table(slide, x, y, w, col_w, rows, row_h, header=True, first_col_dark=False, size=13):
    nr, nc = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(row_h * nr))
    t = gt.table
    # şablon stilini kapat
    tblPr = t._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    for i, cw in enumerate(col_w):
        t.columns[i].width = Inches(cw)
    for r in range(nr):
        t.rows[r].height = Inches(row_h)
        for c in range(nc):
            cell = t.cell(r, c)
            is_head = header and r == 0
            is_key = first_col_dark and c == 0
            cell.fill.solid()
            cell.fill.fore_color.rgb = DARK if is_head else (SLATE if is_key else (LIGHT if r % 2 else WHITE))
            cell.margin_left = cell.margin_right = Inches(0.12)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            tf.text = ""
            run = tf.paragraphs[0].add_run()
            run.text = rows[r][c]
            run.font.name = BODY
            run.font.size = Pt(size)
            run.font.bold = is_head or is_key
            run.font.color.rgb = WHITE if (is_head or is_key) else (DARK if not rows[r][c].startswith("[") else MUTED)
            # ince gri kenarlıklar
            tcPr = cell._tc.get_or_add_tcPr()
            for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = etree.SubElement(tcPr, qn(tag), w=str(Pt(0.75)))
                sf = etree.SubElement(ln, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:srgbClr"), val="C9D1D8")
            # kenarlık elemanları solidFill'den önce gelmeli
            fill_el = tcPr.find(qn("a:solidFill"))
            if fill_el is not None:
                tcPr.remove(fill_el)
                tcPr.append(fill_el)
    return t


# =========================================================
# 0. KAPAK
# =========================================================
s = prs.slides.add_slide(BLANK)
s.background.fill.solid()
s.background.fill.fore_color.rgb = DARK
tag = box(s, 0.8, 0.8, 3.6, 0.45, fill=ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
text(s, 0, 0, 0, 0, "HATA · RW · OPERASYON RAPORU", size=12, bold=True, color=WHITE, font=HEAD,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=tag)
text(s, 0.8, 1.7, 11.7, 1.5, "[Hata / Problem Başlığı]", size=44, bold=True, color=WHITE, font=HEAD,
     anchor=MSO_ANCHOR.BOTTOM)
text(s, 0.8, 3.3, 11.7, 0.5, "Kök Neden Analizi ve Düzeltici Faaliyet Raporu", size=20, color=ORANGE)

fields = [("HAZIRLAYAN", "[Ad Soyad – Proses Mühendisi]"), ("DEPARTMAN / HAT", "[Bölüm / Hat]"),
          ("TARİH", "[GG.AA.YYYY]"), ("RAPOR NO", "[RW-YYYY-NNN]"),
          ("PARÇA NO / İŞ EMRİ", "[Parça No / İş Emri]"), ("OPERASYON", "[Op No – Tanım]")]
cw, gx = 3.7, 0.3
for i, (lab, val) in enumerate(fields):
    cx = 0.8 + (i % 3) * (cw + gx)
    cy = 4.45 + (i // 3) * 1.05
    card = box(s, cx, cy, cw, 0.85, fill=SLATE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, cx + 0.2, cy + 0.12, cw - 0.4, 0.25, lab, size=10, bold=True, color=ORANGE, font=HEAD)
    text(s, cx + 0.2, cy + 0.4, cw - 0.4, 0.35, val, size=15, color=WHITE)
s.notes_slide.notes_text_frame.text = (
    "Köşeli parantez içindeki tüm alanları doldurun. Başlık kısa ve hatayı tarif eden bir cümle olmalı "
    "(ör. 'Op 40 delik çapı tolerans dışı'). Rapor no birim içi takip numarasıdır.")

# =========================================================
# 1. YÖNETİCİ ÖZETİ
# =========================================================
s = content_slide(1, "Yönetici Özeti",
                  "Raporun tamamını tek sayfada anlatın: her kutuya 1–3 kısa cümle yeterli.",
                  "Bu sayfa en son doldurulur. Toplantıda yalnızca bu sayfa gösterilse bile konu anlaşılabilmeli.")
cards = [("Problem", "!", RED, "[Ne oldu, nerede, kaç adet etkilendi?]"),
         ("Kök Neden", "?", ORANGE, "[Doğrulanmış kök neden – tek cümle]"),
         ("Çözüm", "✓", SLATE, "[Alınan kalıcı aksiyon(lar)]"),
         ("Durum ve Sonuç", "●", GREEN, None)]
W2 = (12.13 - 0.3) / 2
for i, (h, g, col, body) in enumerate(cards):
    cx = 0.6 + (i % 2) * (W2 + 0.3)
    cy = 1.75 + (i // 2) * (2.35 + 0.3)
    box(s, cx, cy, W2, 2.35, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    icon_circle(s, cx + 0.3, cy + 0.3, 0.5, g, col)
    text(s, cx + 0.95, cy + 0.3, W2 - 1.2, 0.5, h, size=18, bold=True, font=HEAD, anchor=MSO_ANCHOR.MIDDLE)
    if body:
        text(s, cx + 0.3, cy + 1.0, W2 - 0.6, 1.15, body, size=15, color=MUTED)
    else:
        text(s, cx + 0.3, cy + 1.0, W2 - 0.6, 1.15,
             ["Durum:   ☐ Açık     ☐ Devam ediyor     ☐ Kapandı",
              [("Etkinlik: ", {"bold": True, "color": DARK}), ("[Doğrulama sonucu]", {})]],
             size=15, color=MUTED, space_after=6)

# =========================================================
# 2. PROBLEM TANIMI (5N1K)
# =========================================================
s = content_slide(2, "Problem Tanımı (5N1K)",
                  "Hatayı yoruma yer bırakmadan tarif edin: veriye ve gözleme dayanın, tahmin yazmayın.",
                  "Ne / Nerede / Ne zaman / Kim / Nasıl / Ne kadar sorularının hepsini yanıtlayın. "
                  "Sağdaki alana hatalı parçanın fotoğrafını veya çizim kesitini ekleyin.")
rows = [["Ne oldu?", "[Hatanın tanımı]"],
        ["Nerede?", "[Hat / tezgâh / operasyon]"],
        ["Ne zaman?", "[Tarih – vardiya]"],
        ["Kim tespit etti?", "[Kişi / birim]"],
        ["Nasıl tespit edildi?", "[Ölçüm, görsel kontrol, sonraki op…]"],
        ["Ne kadar?", "[Etkilenen adet / oran]"]]
table(s, 0.6, 1.75, 7.4, [2.4, 5.0], rows, 0.8, header=False, first_col_dark=True, size=14)
photo_placeholder(s, 8.3, 1.75, 4.43, 3.3, "Hatalı parça fotoğrafı\nveya çizim kesiti")
box(s, 8.3, 5.3, 4.43, 1.25, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
text(s, 8.5, 5.42, 4.03, 1.0,
     [[("Spesifikasyon:  ", {"bold": True, "color": DARK}), ("[Ölçü ± tolerans]", {})],
      [("Gerçekleşen:    ", {"bold": True, "color": DARK}), ("[Ölçülen değer]", {})]],
     size=14, color=MUTED, anchor=MSO_ANCHOR.MIDDLE, space_after=6)

# =========================================================
# 3. ETKİ VE KAPSAM
# =========================================================
s = content_slide(3, "Etki ve Kapsam",
                  "Hatanın büyüklüğünü sayılarla gösterin; uygun olan hata türünü işaretleyin.",
                  "Maliyet etkisi bilinmiyorsa tahmini olduğunu belirtin. Hata türünde birden fazla seçim yapılabilir.")
stats = [("Etkilenen parça", "adet"), ("RW / hurda", "adet"), ("Kayıp süre", "saat"), ("Maliyet etkisi", "₺")]
W4 = (12.13 - 3 * 0.3) / 4
for i, (lab, unit) in enumerate(stats):
    cx = 0.6 + i * (W4 + 0.3)
    box(s, cx, 1.75, W4, 1.9, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    text(s, cx + 0.3, 1.95, W4 - 0.6, 0.9, "[ — ]", size=40, bold=True, color=ORANGE, font=HEAD)
    text(s, cx + 0.3, 2.9, W4 - 0.6, 0.5, [[(lab, {"bold": True, "color": DARK}), (f"  ({unit})", {})]],
         size=14, color=MUTED)
text(s, 0.6, 3.95, 6, 0.35, "HATA TÜRÜ", size=12, bold=True, color=SLATE, font=HEAD)
chips = ["Parça hatası", "Rework (RW)", "Operasyon hatası", "Ölçüm hatası", "Malzeme / tedarik", "Diğer"]
W6 = (12.13 - 5 * 0.2) / 6
for i, c in enumerate(chips):
    b = box(s, 0.6 + i * (W6 + 0.2), 4.35, W6, 0.5, fill=WHITE, line=SLATE,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    text(s, 0, 0, 0, 0, "☐  " + c, size=13, color=DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=b)
box(s, 0.6, 5.15, 12.13, 1.5, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
text(s, 0.85, 5.3, 11.6, 0.35, "Sonraki operasyona / müşteriye etkisi", size=14, bold=True, font=HEAD)
text(s, 0.85, 5.7, 11.6, 0.8, "[Teslimat gecikmesi, sonraki operasyonda duruş, müşteri şikâyeti vb.]",
     size=14, color=MUTED)

# =========================================================
# 4. ANLIK ÖNLEMLER
# =========================================================
s = content_slide(4, "Anlık Önlemler",
                  "Hata büyümeden alınan ilk tedbirler: şüpheli parçalar ayrıldı mı, hat güvene alındı mı?",
                  "Bu adımlar kök neden bulunmadan önce, hatanın yayılmasını durdurmak için yapılan işlerdir.")
cols = [("A", "Ayrıştırma ve karantina", "[Şüpheli parçalar nereye ayrıldı, nasıl etiketlendi?]"),
        ("B", "Stok ve hattaki parçaların kontrolü", "[Hangi parti / seri kontrol edildi, sonuç?]"),
        ("C", "İlgili birimlerin bilgilendirilmesi", "[Kalite, planlama, sonraki op, müşteri…]")]
W3 = (12.13 - 2 * 0.3) / 3
for i, (g, h, body) in enumerate(cols):
    cx = 0.6 + i * (W3 + 0.3)
    box(s, cx, 1.75, W3, 4.9, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    icon_circle(s, cx + 0.3, 2.0, 0.55, g, SLATE)
    text(s, cx + 0.3, 2.75, W3 - 0.6, 0.75, h, size=17, bold=True, font=HEAD)
    text(s, cx + 0.3, 3.75, W3 - 0.6, 2.8,
         [[("Yapılan", {"bold": True, "color": DARK, "size": 12})], body,
          [("Adet / kapsam", {"bold": True, "color": DARK, "size": 12})], "[ … ]",
          [("Sorumlu / tarih", {"bold": True, "color": DARK, "size": 12})], "[Ad Soyad – GG.AA]"],
         size=14, color=MUTED, space_after=7)

# =========================================================
# 5. BALIK KILÇIĞI
# =========================================================
s = content_slide(5, "Kök Neden Analizi – Balık Kılçığı",
                  "Her başlık için olası nedenleri yazın; doğrulanan kök nedeni turuncu kutuya alın.",
                  "6M: İnsan, Makine, Metot, Malzeme, Ölçüm, Çevre. Olası nedenleri ekip ile beyin fırtınası "
                  "yaparak çıkarın, ardından veri / deneme ile doğrulayın. Elenen nedenlerin üzerini çizebilirsiniz.")
SPINE_Y = 4.2
ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(0.6), Inches(SPINE_Y), Inches(10.7), Inches(SPINE_Y))
ln.line.color.rgb = SLATE
ln.line.width = Pt(3)
lnEl = ln.line._get_or_add_ln()
tail = etree.SubElement(lnEl, qn("a:tailEnd"))
tail.set("type", "triangle")
head = box(s, 10.75, 3.55, 1.98, 1.3, fill=RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
text(s, 0, 0, 0, 0, [[("PROBLEM", {"bold": True, "size": 12, "font": HEAD})], "[Kısa hata tanımı]"],
     size=13, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=head)
cats = ["İnsan", "Makine", "Metot", "Malzeme", "Ölçüm", "Çevre"]
for i, cat in enumerate(cats):
    top = i < 3
    cx = 0.8 + (i % 3) * 3.3
    hy = 1.75 if top else 6.2
    if top:
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx + 0.15), Inches(hy + 0.42),
                                   Inches(cx + 0.85), Inches(SPINE_Y))
    else:
        c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(cx + 0.85), Inches(SPINE_Y),
                                   Inches(cx + 0.15), Inches(hy))
    c.line.color.rgb = SLATE
    c.line.width = Pt(1.75)
    hb = box(s, cx, hy, 1.9, 0.42, fill=SLATE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    text(s, 0, 0, 0, 0, cat, size=14, bold=True, color=WHITE, font=HEAD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=hb)
    ty = 2.3 if top else 4.85
    bullets(s, cx + 0.7, ty, 2.45, 1.2, ["[olası neden]", "[olası neden]"], size=14, color=MUTED)
# kök neden etiketi örneği
rc = box(s, 10.75, 1.75, 1.98, 1.4, fill=WHITE, line=ORANGE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08, lw=2.5)
text(s, 0, 0, 0, 0, [[("KÖK NEDEN", {"bold": True, "size": 12, "color": ORANGE, "font": HEAD})],
                     "[Doğrulanan neden]"], size=13, color=DARK, align=PP_ALIGN.CENTER,
     anchor=MSO_ANCHOR.MIDDLE, shape=rc)

# =========================================================
# 6. 5 NEDEN
# =========================================================
s = content_slide(6, "Kök Neden Analizi – 5 Neden",
                  "Problemden başlayın; kök nedene ulaşana kadar her cevaba yeniden 'Neden?' diye sorun.",
                  "5 sayısı kesin değildir; kök nedene 3. veya 4. soruda ulaşılabilir. Son cevap, ortadan "
                  "kaldırıldığında hatanın tekrar etmeyeceği neden olmalıdır.")
for i in range(5):
    ry = 1.75 + i * 0.97
    chip = box(s, 0.6, ry, 1.45, 0.75, fill=ORANGE if i == 4 else SLATE, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
               radius=0.12)
    text(s, 0, 0, 0, 0, f"{i + 1}. NEDEN", size=14, bold=True, color=WHITE, font=HEAD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=chip)
    fb = box(s, 2.2, ry, 6.35, 0.75, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
    q = "[Problem neden oluştu?]" if i == 0 else f"[{i}. cevap neden oluştu?]"
    text(s, 0, 0, 0, 0, q, size=14, color=MUTED, anchor=MSO_ANCHOR.MIDDLE, margin=0.2, shape=fb)
rc = box(s, 8.85, 1.75, 3.88, 4.63, fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
text(s, 9.15, 2.0, 3.3, 0.35, "KÖK NEDEN", size=13, bold=True, color=ORANGE, font=HEAD)
text(s, 9.15, 2.4, 3.3, 1.3, "[Doğrulanmış kök neden ifadesi]", size=20, bold=True, color=WHITE, font=HEAD)
text(s, 9.15, 3.85, 3.3, 2.4,
     [[("Kategori (6M)", {"bold": True, "color": ORANGE, "size": 12})], "[İnsan / Makine / Metot / …]",
      [("Nasıl doğrulandı?", {"bold": True, "color": ORANGE, "size": 12})], "[Deneme, ölçüm, veri analizi…]"],
     size=14, color=LIGHT, space_after=7)

# =========================================================
# 7. AKSİYON PLANI
# =========================================================
s = content_slide(7, "Çözüm ve Aksiyon Planı",
                  "Her aksiyonun tek bir sorumlusu ve net bir termini olmalı.",
                  "Tür: D = Düzeltici (bu hatayı gideren), Ö = Önleyici (tekrarını engelleyen). "
                  "Gerekirse tabloya satır ekleyin.")
rows = [["#", "Aksiyon", "Tür (D / Ö)", "Sorumlu", "Termin", "Durum"]]
for i in range(5):
    rows.append([str(i + 1), "[Aksiyon tanımı]", "[D / Ö]", "[Ad Soyad]", "[GG.AA]", "[Durum]"])
table(s, 0.6, 1.75, 12.13, [0.55, 5.2, 1.5, 2.0, 1.33, 1.55], rows, 0.6, size=14)
text(s, 0.6, 5.75, 3.0, 0.4, "DURUM GÖSTERİMİ", size=12, bold=True, color=SLATE, font=HEAD,
     anchor=MSO_ANCHOR.MIDDLE)
for i, (lab, col) in enumerate([("Açık", RED), ("Devam ediyor", AMBER), ("Tamamlandı", GREEN)]):
    lx = 3.0 + i * 2.3
    box(s, lx, 5.85, 0.22, 0.22, fill=col, shape=MSO_SHAPE.OVAL)
    text(s, lx + 0.35, 5.75, 1.8, 0.4, lab, size=14, anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.6, 6.2, 12.0, 0.4, "D = Düzeltici aksiyon (hatayı giderir)   ·   Ö = Önleyici aksiyon (tekrarını engeller)",
     size=12, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# 8. STANDARTLAŞTIRMA
# =========================================================
s = content_slide(8, "Standartlaştırma ve Yaygınlaştırma",
                  "Çözümün kalıcı olması için güncellenen dokümanları ve benzer yerlere aktarımı işaretleyin.",
                  "Güncellenmesi gerekmeyen kalemler için 'Gerekli değil' işaretleyin. Doküman no ve revizyonu yazın.")
docs = ["Proses / iş talimatı", "Kontrol planı", "PFMEA", "CNC programı / fikstür",
        "Operatör eğitimi", "Hata önleme (poka-yoke)"]
GW, GH = (8.0 - 0.3) / 2, (4.9 - 0.6) / 3
for i, d in enumerate(docs):
    cx = 0.6 + (i % 2) * (GW + 0.3)
    cy = 1.75 + (i // 2) * (GH + 0.3)
    box(s, cx, cy, GW, GH, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, cx + 0.25, cy + 0.15, GW - 0.5, 0.4, d, size=16, bold=True, font=HEAD)
    text(s, cx + 0.25, cy + 0.58, GW - 0.5, 0.75,
         ["☐ Güncellendi      ☐ Gerekli değil", "[Doküman no / revizyon]"],
         size=13, color=MUTED, space_after=2)
yb = box(s, 8.9, 1.75, 3.83, 4.9, fill=SLATE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
text(s, 9.2, 2.0, 3.25, 0.4, "YAYGINLAŞTIRMA", size=13, bold=True, color=ORANGE, font=HEAD)
text(s, 9.2, 2.5, 3.25, 4.0,
     [[("Benzer parça / hat / tezgâhlar", {"bold": True, "color": WHITE})], "[Liste]",
      [("Uygulanan", {"bold": True, "color": WHITE})], "[Ne yapıldı?]",
      [("Tarih", {"bold": True, "color": WHITE})], "[GG.AA.YYYY]"],
     size=14, color=LIGHT, space_after=4)

# =========================================================
# 9. DOĞRULAMA
# =========================================================
s = content_slide(9, "Doğrulama – Önce / Sonra",
                  "Aksiyonun işe yaradığını veriyle gösterin: aynı ölçüt, aynı yöntem, öncesi ve sonrası.",
                  "Fotoğraf, ölçüm grafiği veya SPC çıktısı ekleyebilirsiniz. Etkinlik kontrolü süresini "
                  "(ör. 30 gün / 500 adet) önceden belirleyin.")
for i, (lab, col) in enumerate([("ÖNCE", RED), ("SONRA", GREEN)]):
    cx = 0.6 + i * 6.53
    box(s, cx, 1.75, 5.6, 3.55, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    p = box(s, cx + 0.25, 1.95, 1.3, 0.4, fill=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    text(s, 0, 0, 0, 0, lab, size=13, bold=True, color=WHITE, font=HEAD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=p)
    photo_placeholder(s, cx + 0.25, 2.5, 5.1, 2.0, "Fotoğraf / ölçüm grafiği")
    text(s, cx + 0.25, 4.65, 5.1, 0.45,
         [[("Ölçüt: ", {"bold": True, "color": DARK}), ("[Hata oranı / ölçü değeri]", {})]], size=14,
         color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
ar = box(s, 6.4, 3.2, 0.53, 0.65, fill=ORANGE, shape=MSO_SHAPE.CHEVRON)
box(s, 0.6, 5.55, 12.13, 1.1, fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
text(s, 0.9, 5.55, 3.2, 1.1, "ETKİNLİK KONTROLÜ", size=13, bold=True, color=ORANGE, font=HEAD,
     anchor=MSO_ANCHOR.MIDDLE)
text(s, 4.0, 5.55, 5.0, 1.1, [[("[ N ]", {"bold": True, "size": 24, "color": WHITE, "font": HEAD}),
                               ("  adet  /  ", {}), ("[ N ]", {"bold": True, "size": 24, "color": WHITE, "font": HEAD}),
                               ("  gün tekrar yok", {})]], size=14, color=LIGHT, anchor=MSO_ANCHOR.MIDDLE)
text(s, 9.1, 5.55, 3.4, 1.1, ["Doğrulayan: [Ad Soyad]", "Tarih: [GG.AA.YYYY]"], size=13, color=LIGHT,
     anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# 10. KAPANIŞ
# =========================================================
s = content_slide(10, "Sonuç, Öğrenilenler ve Onay",
                  "Bu olaydan çıkarılan dersleri yazın ve raporu onaya sunun.",
                  "Öğrenilen dersler başka projelerde de kullanılabilecek genel çıkarımlar olmalı. "
                  "Rapor, onaylayanın imzası ile kapanır.")
box(s, 0.6, 1.75, 5.9, 4.9, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
icon_circle(s, 0.9, 2.0, 0.5, "★", ORANGE)
text(s, 1.55, 2.0, 4.7, 0.5, "Öğrenilen Dersler", size=18, bold=True, font=HEAD, anchor=MSO_ANCHOR.MIDDLE)
bullets(s, 0.9, 2.8, 5.3, 3.6, ["[Ders 1]", "[Ders 2]", "[Ders 3]"], size=15, color=MUTED)
text(s, 6.9, 1.75, 5.83, 0.4, "RAPOR DURUMU", size=12, bold=True, color=SLATE, font=HEAD)
for i, (lab, col) in enumerate([("☐  AÇIK", RED), ("☐  KAPANDI", GREEN)]):
    b = box(s, 6.9 + i * 3.0, 2.2, 2.8, 0.55, fill=WHITE, line=col, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
            radius=0.5, lw=2)
    text(s, 0, 0, 0, 0, lab, size=15, bold=True, color=col, font=HEAD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=b)
text(s, 6.9, 3.1, 5.83, 0.4, "ONAY", size=12, bold=True, color=SLATE, font=HEAD)
rows = [["Rol", "Ad Soyad", "Tarih", "İmza"],
        ["Hazırlayan\n(Proses Müh.)", "[ ]", "[ ]", ""],
        ["Kontrol eden", "[ ]", "[ ]", ""],
        ["Onaylayan\n(Kalite)", "[ ]", "[ ]", ""]]
table(s, 6.9, 3.5, 5.83, [1.75, 1.7, 1.08, 1.3], rows, 0.78, size=13)

prs.save(OUT)
print("kaydedildi:", OUT)
