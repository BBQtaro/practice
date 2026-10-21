# -*- coding: utf-8 -*-
"""積算根拠図（PDF）を作る: 施工図一式 PDF の該当ページに拾い出し範囲を重ね描きする

python3 build_takeoff_drawings.py <施工図一式.pdf>  ->  GU富士見_2413_積算根拠図.pdf
"""
import sys
import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont
import takeoff_geometry as g

SRC = sys.argv[1]
OUT = "GU富士見_2413_積算根拠図.pdf"
S = 2.0                       # 描画倍率（1 = 72dpi）
K = S / 1.2
X0, Y0 = 414 * K, 154 * K     # 通り芯 X21 / Y1 の画素位置（A03〜A05 共通）
PPM = 144.75 * K / 8500       # px / mm
FONT = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"


def font(size):
    return ImageFont.truetype(FONT, int(size * S / 2))


def P(pt):
    return (X0 + pt[0] * PPM, Y0 + pt[1] * PPM)


def hex2rgba(h, a):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)) + (a,)


pdf = pdfium.PdfDocument(SRC)


def page(n):
    return pdf[n - 1].render(scale=S).to_pil().convert("RGBA")


def header(img, title, sub):
    d = ImageDraw.Draw(img)
    d.rectangle([20, 14, 1250 * S / 2, 92 * S / 2], fill=(255, 255, 255, 235), outline=(200, 30, 30, 255), width=3)
    d.text((34, 22), "積算根拠図 ｜ " + title, font=font(30), fill=(200, 30, 30, 255))
    d.text((34, 22 + 40 * S / 2), sub, font=font(17), fill=(40, 40, 40, 255))


def legend(img, items, x, y, title="凡例"):
    d = ImageDraw.Draw(img)
    h = (len(items) + 1) * 24 * S / 2 + 16
    w = 540 * S / 2
    d.rectangle([x, y, x + w, y + h], fill=(255, 255, 255, 240), outline=(60, 60, 60, 255), width=2)
    d.text((x + 10, y + 6), title, font=font(17), fill=(0, 0, 0, 255))
    for i, (col, txt) in enumerate(items):
        yy = y + 8 + (i + 1) * 24 * S / 2
        d.rectangle([x + 10, yy, x + 40, yy + 16 * S / 2], fill=hex2rgba(col, 200))
        d.text((x + 50, yy - 2), txt, font=font(14), fill=(0, 0, 0, 255))


pages = []

# ---- 0. 表紙（数量集計）
W_, H_ = int(1191 * S), int(842 * S)
cover = Image.new("RGBA", (W_, H_), (255, 255, 255, 255))
d = ImageDraw.Draw(cover)
d.text((80, 60), "GU ららぽーと富士見店（2413区画）積算根拠図", font=font(48), fill=(0, 0, 0, 255))
d.text((80, 60 + 70 * S / 2), "元図：施工図一式（A03 平面図・A04 床伏図・A05 天井伏図・照明器具表・A16〜A21 展開図）／縮尺 1/200 を通り芯基準で 1m グリッド読み取り",
       font=font(17), fill=(60, 60, 60, 255))
y = 60 + 130 * S / 2
d.text((80, y), "■ 図面拾い 数量集計", font=font(26), fill=(0, 0, 0, 255))
y += 40 * S / 2
for k, v in g.summary().items():
    d.text((100, y), f"{k}", font=font(18), fill=(0, 0, 0, 255))
    d.text((100 + 560 * S / 2, y), f"{v}", font=font(18), fill=(0, 0, 160, 255))
    y += 28 * S / 2
y += 14 * S / 2
d.text((80, y), "■ 壁仕上げ（面積・展開図による面別）", font=font(26), fill=(0, 0, 0, 255))
y += 40 * S / 2
for k, v in g.finish_totals().items():
    d.text((100, y), k, font=font(18), fill=(0, 0, 0, 255))
    d.text((100 + 560 * S / 2, y), f"{v} ㎡", font=font(18), fill=(0, 0, 160, 255))
    y += 28 * S / 2
x2 = W_ / 2 + 40
y2 = 60 + 130 * S / 2
d.text((x2, y2), "■ 収録ページ", font=font(26), fill=(0, 0, 0, 255))
for i, t in enumerate(["1. 床仕上げ範囲（A04 床伏図）", "2. 壁 下地・仕上げ範囲（A03 平面図）", "3. 天井仕上げ範囲（A05 天井伏図）",
                       "4. 照明器具 台数（照明配置図）", "5. コンセント・弱電 台数（コンセント配置図）", "6〜9. 展開図 ↔ 壁ID 対応（A16〜A19）"]):
    d.text((x2 + 20, y2 + (i + 1) * 30 * S / 2 + 10), t, font=font(18), fill=(0, 0, 0, 255))
notes = ["・室面積は平面図の面積表（正式値）で計上。図上概算は確認用（差 ±10% 程度）。",
         "・壁長さは通り芯からの図上読み取り。高さは CH3,750（SW は 3,155、下がり壁 650）。",
         "・器具・コンセントは照明器具表の総数量で計上（配置図と目視照合）。",
         "・GU 支給品（照明器具・床材・什器・サイン）は施工費のみ。",
         "・見積 v2（2021実績数量ベース）との差は 確認用HTML の比較表を参照。"]
for i, t in enumerate(notes):
    d.text((x2, y2 + 260 * S / 2 + i * 28 * S / 2), t, font=font(16), fill=(80, 80, 80, 255))
pages.append(cover)

# ---- 1. 床仕上げ（A04 床伏図 = p9）
img = page(9)
ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
d = ImageDraw.Draw(ov)
for z in g.zone_rows():
    d.polygon([P(p) for p in z["pts"]], fill=hex2rgba(z["color"], 95), outline=hex2rgba(z["color"], 255))
for _, name, pts in g.EXCLUDED:
    d.polygon([P(p) for p in pts], fill=(120, 120, 120, 110))
img = Image.alpha_composite(img, ov)
d = ImageDraw.Draw(img)
for z in g.zone_rows():
    cx = sum(p[0] for p in z["pts"]) / len(z["pts"])
    cy = sum(p[1] for p in z["pts"]) / len(z["pts"])
    x, y = P((cx, cy))
    txt = f'{z["id"]} {z["name"]}\n正式 {z["official"]}㎡ / 図上 {z["area"]}㎡'
    d.multiline_text((x - 60 * S / 2, y - 14), txt, font=font(13), fill=(150, 0, 0, 255),
                     stroke_width=3, stroke_fill=(255, 255, 255, 255))
header(img, "1. 床仕上げ範囲", "A04 床伏図に重ね描き。売場・レジ・FRブース外＝ES3001-45（材料GU支給）、BOH・SW＝P-41、FRブース内＝WBH-603。灰色＝階段室（A工事・対象外）")
legend(img, [(z["color"], f'{z["id"]} {z["name"]}：{z["finish"]}') for z in g.zone_rows()], 40, 900 * S / 2)
pages.append(img)

# ---- 2. 壁（A03 平面図 = p8）
img = page(8)
d = ImageDraw.Draw(img)
for w in g.wall_rows():
    pts = [P(p) for p in w["pts"]]
    col = hex2rgba(w["color"], 230)
    if w["id"] in g.SEGMENT_WALLS:
        for a, b in zip(pts[::2], pts[1::2]):
            d.line([a, b], fill=col, width=int(9 * S / 2))
        lx, ly = pts[0]
    else:
        d.line(pts, fill=col, width=int(9 * S / 2), joint="curve")
        mid = len(pts) // 2
        lx, ly = ((pts[mid - 1][0] + pts[mid][0]) / 2, (pts[mid - 1][1] + pts[mid][1]) / 2) if len(pts) > 1 else pts[0]
    d.text((lx + 8, ly + 4), f'{w["id"]} L={w["length"]}m', font=font(14), fill=(200, 0, 0, 255),
           stroke_width=3, stroke_fill=(255, 255, 255, 255))
for cx, cy in g.COLUMNS:
    x, y = P((g.GRID_X[cx], g.GRID_Y[cy]))
    r = g.COLUMN_SIZE / 2 * g.PPM if hasattr(g, "PPM") else 600 * PPM
    d.rectangle([x - r, y - r, x + r, y + r], outline=(0, 160, 120, 255), width=5)
header(img, "2. 壁 下地・仕上げ範囲", "A03 平面図に重ね描き。線色＝仕上げ（青 AEP N-87／黒 黒クロス／橙 白クロス／茶 化粧PB／灰 素地・A工事壁／桃 SW／緑 下がり壁）。緑枠＝柱巻きミラー対象柱 7本")
legend(img, [("#2b6cb0", "AEP N-87 3分ツヤ（売場・補正室・SW）"), ("#111111", "黒クロス LW-4207（レジ背面・右壁）"),
             ("#c05621", "白クロス LW-4630/4635（試着室）"), ("#975a16", "化粧PB 木目・桐（休憩室・店長室）"),
             ("#a0aec0", "素地＋ワイド巾木H300（倉庫）"), ("#6b7280", "A工事壁 直接仕上げ（階段室）"),
             ("#d53f8c", "ショーウィンドウ背面（CH3,155）"), ("#38a169", "店頭 下がり壁 H650")], 40, 980 * S / 2, "凡例（壁）")
pages.append(img)

# ---- 3. 天井（A05 天井伏図 = p10）
img = page(10)
ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
d = ImageDraw.Draw(ov)
for z in g.zone_rows():
    col = "#e0a800" if z["name"] == "店長室" else "#3aa0a0"
    d.polygon([P(p) for p in z["pts"]], fill=hex2rgba(col, 80), outline=hex2rgba(col, 255))
img = Image.alpha_composite(img, ov)
header(img, "3. 天井仕上げ範囲", f'A05 天井伏図。全室 既存A工事天井転用＋AEP N-95（{g.summary()["天井 AEP（全室, ㎡）"]}㎡）。店長室のみ 天井PB不燃12.5 貼替＋AEP。SW は CH3,155')
legend(img, [("#3aa0a0", "既存天井転用＋AEP N-95 ツヤ消し"), ("#e0a800", "店長室：PB12.5 貼替＋AEP")], 40, 1000 * S / 2, "凡例（天井）")
pages.append(img)


# ---- 4/5. 照明・コンセント（器具表の台数をスタンプ）
def stamp_page(n, title, sub, rows):
    img = page(n)
    header(img, title, sub)
    d = ImageDraw.Draw(img)
    x, y = img.width - 760 * S / 2, 120 * S / 2
    h = (len(rows) + 1) * 30 * S / 2 + 20
    d.rectangle([x, y, x + 720 * S / 2, y + h], fill=(255, 255, 230, 240), outline=(200, 30, 30, 255), width=3)
    d.text((x + 14, y + 8), "拾い数量（照明器具表 総数量）", font=font(18), fill=(200, 30, 30, 255))
    for i, (name, q, src) in enumerate(rows):
        d.text((x + 14, y + 12 + (i + 1) * 30 * S / 2), f"{name}：{q}", font=font(17), fill=(0, 0, 0, 255))
    return img


pages.append(stamp_page(11, "4. 照明器具 台数", "照明配置図。器具は GU 支給（シマダ手配）→ 開口・取付・配線・落下防止ワイヤーのみ計上",
                        [r for r in g.FIXTURES if "コンセント" not in r[0] and "アウトレット" not in r[0]]))
pages.append(stamp_page(16, "5. コンセント・弱電 台数", "コンセント配置図。取付・配線（材工）で計上。カットリレーはB工事（空配管・電源のみC工事）",
                        [r for r in g.FIXTURES if "コンセント" in r[0] or "アウトレット" in r[0]]))

# ---- 6〜9. 展開図 ↔ 壁ID
ELEV_MAP = {33: ["W17 店頭下がり壁（展開1・2）", "W01 売場左壁（展開3）"],
            34: ["W07/W08 奥壁・FR前面（展開4）", "W02/W09 レジ 黒クロス（展開5・5'・6・7）"],
            35: ["W05 階段室 A工事壁（展開11）", "W06 右壁 H3664什器（展開13）", "W08/W12/W16 試着室 クロス（展開16〜19）", "W03/W04 備品庫 素地（展開20〜22）"],
            36: ["W09/W13/W14 補正室 AEP（展開24〜27）", "W11〜W14 店長室・休憩室 化粧PB（展開28〜36）"]}
for n, lines in ELEV_MAP.items():
    img = page(n)
    header(img, f"{6 + list(ELEV_MAP).index(n)}. 展開図 ↔ 壁ID 対応", "壁の仕上げ種別・高さはこの展開図で確認（面数は確認用HTMLの壁一覧に記載）")
    d = ImageDraw.Draw(img)
    x, y = img.width - 820 * S / 2, 110 * S / 2
    d.rectangle([x, y, x + 790 * S / 2, y + (len(lines) + 1) * 30 * S / 2 + 16], fill=(255, 255, 230, 240), outline=(200, 30, 30, 255), width=3)
    for i, t in enumerate(lines):
        d.text((x + 14, y + 12 + i * 30 * S / 2), t, font=font(17), fill=(0, 0, 0, 255))
    pages.append(img)

rgb = [p.convert("RGB") for p in pages]
rgb[0].save(OUT, save_all=True, append_images=rgb[1:], resolution=72 * S, quality=70)
# 確認用HTML 用に床・壁ページの縮小画像も保存
for i, name in [(1, "base_floor.jpg"), (2, "base_wall.jpg")]:
    im = rgb[i].copy()
    im.thumbnail((1600, 1600))
    im.save(name, quality=72)
print("saved", OUT, len(rgb), "pages")
