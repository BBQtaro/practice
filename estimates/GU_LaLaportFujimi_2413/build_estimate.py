# -*- coding: utf-8 -*-
"""GU ららぽーと富士見店（2413区画）C工事 積算・見積・工程表 生成スクリプト

python3 build_estimate.py  ->  GU_ららぽーと富士見店_見積書_工程表.xlsx
数量・単価を変更する場合は ITEMS を編集して再実行するか、Excel 上で直接修正する
（金額・小計・合計はすべて数式なので Excel 上の修正も自動で反映される）。
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = "GU_ららぽーと富士見店_見積書_工程表.xlsx"
FONT = "游ゴシック"
ISSUE_DATE = dt.date(2026, 10, 21)
START = dt.date(2027, 1, 12)  # C工事 着工（初夜）仮定

thin = Side(style="thin", color="808080")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FILL = PatternFill("solid", fgColor="DCE6F1")
SEC_FILL = PatternFill("solid", fgColor="F2F2F2")
SUB_FILL = PatternFill("solid", fgColor="FFF2CC")
BAR_FILL = PatternFill("solid", fgColor="4F81BD")
BAR2_FILL = PatternFill("solid", fgColor="F79646")
BAR3_FILL = PatternFill("solid", fgColor="9BBB59")
SUN_FILL = PatternFill("solid", fgColor="FDE9D9")
YEN = '#,##0'


def f(size=10, bold=False, color="000000"):
    return Font(name=FONT, size=size, bold=bold, color=color)


# ---------------------------------------------------------------------------
# 数量根拠（図面より）
# ---------------------------------------------------------------------------
AREA_TOTAL = 1064.0          # 照明器具表 床面積
AREA_BOH = 208.9             # フロアマップ BOH/TOTAL（試着室含む）
AREA_FR = 36.0
AREA_SW = 21.6
AREA_REG = 66.8
BOH_ROOMS = {"補正室": 15.8, "倉庫(STOCK-1)": 94.4, "備品庫(STOCK-2)": 17.9,
             "休憩室": 35.5, "店長室": 9.5}
AREA_BOH_ROOMS = round(sum(BOH_ROOMS.values()), 1)          # 173.1
AREA_SALES = round(AREA_TOTAL - AREA_BOH, 1)               # 855.1（レジ・SW含む）
FACADE_LEN = 20.8            # ファサードボーダー重量計算 6.790+7.283+6.745
CH = 3.75
PERIM_WALL = 110.0           # 区画外周（間口4スパン×8.5m=34m、奥行≒31.3m）からファサード20.8mを除く
PART_LEN = 75.0              # BOH間仕切長さ（平面図より概算）

QTY_BASIS = [
    ("区画面積（照明器具表 記載床面積）", AREA_TOTAL, "㎡", "照明器具表「床面積 1064.0㎡」"),
    ("BOH合計（試着室含む）", AREA_BOH, "㎡", "フロアマップ BOH/TOTAL AREA 208.9㎡"),
    ("　うち 試着室(FR)", AREA_FR, "㎡", "フロアマップ"),
    ("　うち バックルーム5室計", AREA_BOH_ROOMS, "㎡", "補正室15.8+倉庫94.4+備品庫17.9+休憩室35.5+店長室9.5"),
    ("売場（レジ・SW含む）", AREA_SALES, "㎡", "1064.0−208.9"),
    ("　うち レジエリア", AREA_REG, "㎡", "フロアマップ"),
    ("　うち ショーウィンドウ", AREA_SW, "㎡", "フロアマップ"),
    ("天井高 CH", CH * 1000, "mm", "各図 2413区画 CH=3750"),
    ("ファサード長さ", FACADE_LEN, "m", "ファサードボーダー重量計算 6.790+7.283+6.745"),
    ("外周C工事壁 長さ（概算）", PERIM_WALL, "m", "通り芯X17〜X21 8500×4=34m、奥行≒1064/34=31.3m→周長130.6m−ファサード20.8m（要実測）"),
    ("BOH間仕切 長さ（概算）", PART_LEN, "m", "平面図より概算（要拾い直し）"),
    ("ダウンライト BD1〜BD6", 99 + 87 + 12 + 59 + 5, "台", "照明器具表 総数量"),
    ("ユニバーサルDL UD6〜UD8", 37 + 79 + 18, "台", "照明器具表 総数量"),
    ("ベースライト BL1/BL2", 8 + 2 + 3, "台", "照明器具表 総数量"),
    ("スポットライト SP1/SP2", 17 + 11, "台", "照明器具表 総数量"),
    ("ダブルコンセント", 41, "個", "照明器具表 コンセント欄"),
    ("抜け止め接地付ダブルコンセント", 42, "個", "照明器具表 コンセント欄"),
    ("POS/TEL/FAX アウトレット", 20, "個", "照明器具表 コンセント欄"),
    ("テナント分電盤", 2, "面", "盤図 U54836 電灯盤-1（200AT）・電灯盤-2（175AT）"),
]

# ---------------------------------------------------------------------------
# 明細（工種, 名称, 仕様, 数量, 単位, 単価, 根拠）
# ---------------------------------------------------------------------------
DL_BD = {"BD1-35 LED67.1W DL55°": (99, 39000), "BD3-35 LED45.5W DL55°": (87, 34000),
         "BD4-35 LED45.5W DL45°": (12, 34000), "BD5-35 LED38.8W DL55°": (59, 30000),
         "BD6-35 LED17.6W DL55°": (5, 24000)}
DL_UD = {"UD6-35 LED34.9W ユニDL20°": (37, 36000), "UD7-35 LED34.9W ユニDL30°": (79, 36000),
         "UD8-35 LED34.9W ユニDL35°": (18, 36000)}
N_DL = sum(q for q, _ in DL_BD.values()) + sum(q for q, _ in DL_UD.values())
N_LIGHT = N_DL + 13 + 28 + 1

S_SALES_FR = round(AREA_SALES + AREA_FR, 1)     # 売場+FR 天井AEP
WALL_PERIM = round(PERIM_WALL * CH, 1)          # 412.5
WALL_PART = round(PART_LEN * CH * 2, 1)         # 562.5（両面）

ITEMS = [
    ("01 仮設工事", [
        ("墨出し", "床・壁・天井・器具位置", AREA_TOTAL, "㎡", 300, "区画面積"),
        ("養生", "区画内・既存仕上・A/B工事設備・後方通路二次養生", 1, "式", 380000, "施工指針書 10-6"),
        ("仮設電気・工事用照明", "漏電遮断器付仮設盤・投光器（照明の準備なし）", 1, "式", 180000, "施工指針書 10-8"),
        ("内部足場", "ローリングタワー・可搬式作業台 CH3750", 1, "式", 280000, ""),
        ("施工許可看板・掲示板・消火器", "自立型 t6ベニヤ、ABC10型", 1, "式", 60000, "施工指針書 6-9"),
        ("日々清掃・片付け", "毎夜", 40, "夜", 15000, "工程40夜"),
        ("竣工クリーニング", "床・建具・器具・ガラス", AREA_TOTAL, "㎡", 550, ""),
        ("資材搬入運搬費", "4t車以下・北荷捌所 22:00以降", 14, "台", 45000, "施工指針書 搬出入ルール"),
        ("産業廃棄物処分", "分別・搬出・マニフェスト", 8, "台", 85000, "4t車換算"),
    ]),
    ("02 解体・撤去工事", [
        ("店頭サッシ撤去", "既存サッシ撤去・処分", 1, "式", 180000, "平面図 注記②（工事区分 要確認）"),
        ("既存設備・残置物撤去", "既存店舗転用部の不要物撤去", 1, "式", 120000, ""),
    ]),
    ("03 軽量鉄骨下地工事", [
        ("壁下地 LGS65", "C工事壁（店舗間仕切前出し）@303 H3750", WALL_PERIM, "㎡", 4200, f"{PERIM_WALL}m×{CH}m"),
        ("間仕切下地 LGS65", "BOH間仕切 @303 H3750", round(PART_LEN * CH, 1), "㎡", 4200, f"{PART_LEN}m×{CH}m"),
        ("天井下地 LGS19型", "BOH天井 @303", AREA_BOH_ROOMS, "㎡", 3800, "バックルーム5室計"),
        ("ファサードボーダー下地", "H650 LGS+ランナー、LEDサイン固定下地 FB t3", round(FACADE_LEN * 1.9, 1), "㎡", 7500, "重量計算書 PB 13.5+22.9+3.1㎡"),
        ("防煙垂れ壁 下地・ボード", "H500 不燃 両面PB", 25, "m", 9500, "平面図 防煙垂れ壁H=500（A工事分除く・要確認）"),
        ("建具開口補強", "", 11, "箇所", 12000, "建具表"),
        ("壁面什器取付下地", "合板t12 増し貼り H2400", round(PERIM_WALL * 2.4, 1), "㎡", 3200, "壁面什器範囲"),
        ("ダウンライト開口・補強", "A工事天井 開口φ150+補強（A-00h）", N_DL, "箇所", 2500, "BD+UD台数"),
    ]),
    ("04 ボード工事", [
        ("壁 PB12.5 張り", "外周C工事壁 片面", WALL_PERIM, "㎡", 1600, ""),
        ("間仕切 PB12.5 張り", "両面", WALL_PART, "㎡", 1600, ""),
        ("天井 化粧石膏ボード GB-D(桐)", "吉野石膏 9.5 不燃NM-0128", AREA_BOH_ROOMS, "㎡", 2300, "仕上表 BOH天井"),
        ("ファサード PB 張り", "", round(FACADE_LEN * 1.9, 1), "㎡", 1700, ""),
        ("天井点検口", "450角（C工事天井部）", 6, "箇所", 15000, ""),
    ]),
    ("05 床工事", [
        ("床下地調整", "ケレン・パテしごき", AREA_TOTAL, "㎡", 900, "既存躯体+水性ウレタン"),
        ("塩ビタイル ES3001-45(グレー)", "川島セルコン 馬貼り FL+3 売場・FRブース外", AREA_SALES, "㎡", 4800, "売場−SW+FRブース外"),
        ("塩ビタイル クラシックP-41(グレー)", "田島 FL+2 BOH・SW", round(AREA_BOH_ROOMS + AREA_SW, 1), "㎡", 3600, "BOH5室+SW"),
        ("塩ビタイル ウッドライン WBH-603", "田島 乱貼り FRブース内", round(AREA_FR * 0.4, 1), "㎡", 5500, "FR面積×40%"),
        ("ソフト巾木 HB-89 H100", "タジマ ライトグレー 売場", 120, "m", 1100, ""),
        ("ソフト巾木 TH75-96 H75", "東リ BOH", 160, "m", 900, ""),
        ("巾木 WTH300-96 H300", "東リ FR", 40, "m", 2600, ""),
        ("床見切り・リースライン際左官調整", "店頭断面詳細図", FACADE_LEN, "m", 3500, ""),
        ("ドットシール", "避難導線・シャッターライン 450P φ30 リンテック", 150, "m", 600, ""),
        ("分電盤まわりイエローテープ", "3M PM1981", 1, "式", 15000, ""),
    ]),
    ("06 塗装工事", [
        ("天井 AEP N-93(白)ツヤ消し", "売場・FR CH3750", S_SALES_FR, "㎡", 1400, "売場+FR"),
        ("壁 AEP N-95(白)ツヤ消し", "売場外周・BOH", round(WALL_PERIM + WALL_PART, 1), "㎡", 1200, ""),
        ("ショーウィンドウ壁 AEP N-87 3分ツヤ", "", 40, "㎡", 1500, ""),
        ("A工事避難扉 SOP N-87", "原状復旧対象", 3, "枚", 18000, "平面図 注記"),
    ]),
    ("07 クロス工事", [
        ("クロス LW-4630(白)", "リリカラ 不燃NM-3991 FRブース・壁面上部", 150, "㎡", 1300, ""),
    ]),
    ("08 木製建具・鋼製建具工事", [
        ("WD-1A〜1D 片開き木製扉", "W800(1D:W1025)×H2125 ポリ化粧合板・三方枠 金物込", 4, "枚", 165000, "建具表"),
        ("WD-3A 親親開き木製扉", "W1250 倉庫 非常開装置付本締錠 DA-1E", 1, "組", 260000, "建具表"),
        ("WD-4A FRブース扉", "木目", 1, "枚", 110000, "建具表"),
        ("WD-5A 金庫隠し扉", "W850 ケースロック", 1, "枚", 120000, "建具表"),
        ("WD-5D 金庫室扉", "W665 固定有", 1, "組", 150000, "建具表 E-06"),
        ("WD-6 フレームレス扉", "隠し丁番・プッシュツマミ", 1, "枚", 95000, "建具表"),
        ("SD-1A 片開きスチールドア", "W800 t1.6 ウレタン塗装 休憩室〜店長室", 1, "枚", 230000, "建具表"),
        ("SG-1 排煙水平ガラリ", "スチール t1.6 白塗装", 3, "箇所", 85000, "建具・ガラリ表（箇所数要確認）"),
        ("鍵（共通キー・マスターキー対応）", "金庫室×5本・共通×5本", 1, "式", 60000, "建具表 注記"),
    ]),
    ("09 金属工事", [
        ("各見切り", "SUS・アルミ見切り（B-07）", 120, "m", 1800, ""),
        ("自動ドアガード", "B-13", 1, "式", 120000, ""),
        ("防犯ネット取付", "A-15", 1, "式", 80000, ""),
        ("台車ガード", "アルポリ t2 ストック扉", 1, "式", 40000, ""),
        ("金庫固定金具", "C-05", 1, "式", 35000, ""),
        ("ピクチャーレール", "VPテーブル上 B-09a", 20, "m", 4500, ""),
    ]),
    ("10 造作家具工事", [
        ("レジバック造作", "C-03", 1, "式", 850000, ""),
        ("レジバック棚板", "W600/W900 C-04", 1, "式", 120000, ""),
        ("試着室（FR）造作", "ブース間仕切・三方枠・カーテンレール D-01〜D-08", 10, "室", 180000, "ブース数は仮定（要確認）"),
        ("FRソファー（大型）", "D-06", 2, "台", 120000, ""),
        ("FR収納棚", "", 1, "式", 250000, ""),
        ("FRカーテン", "単品図", 10, "枚", 22000, ""),
        ("補正室棚板", "F-03", 1, "式", 150000, ""),
        ("ストックルーム棚", "棚2段（平面図）", 1, "式", 600000, "什器支給品の場合は減額"),
        ("休憩室流し台まわり", "F-05（給排水接続は別途）", 1, "式", 280000, ""),
        ("店長室収納棚", "H2164", 1, "式", 180000, ""),
        ("タペストリーボックス", "B-14", 2, "台", 90000, ""),
        ("ファサードショーウィンドー造作", "B-10", 1, "式", 650000, ""),
        ("ファサードボンデパネル", "B-15", 1, "式", 450000, ""),
        ("ロールスクリーン", "ショーウィンドウ B-11", 3, "台", 65000, ""),
    ]),
    ("11 ガラス・ミラー工事", [
        ("柱巻ミラー（独立柱）", "B-03a", 6, "本", 280000, "本数は通り芯から仮定"),
        ("柱巻ミラー（什器隣接柱）", "B-03b", 4, "本", 160000, ""),
        ("FRミラー", "各ブース", 10, "枚", 28000, ""),
    ]),
    ("12 サイン工事", [
        ("ファサード内照式ロゴサイン 取付", "G-10【支給品】電源接続含む", 1, "式", 180000, ""),
        ("GUロゴアクリルサイン 製作取付", "G-03", 1, "式", 120000, ""),
        ("ガラスサッシロゴサイン", "G-05 シート", 1, "式", 60000, ""),
        ("欄間サイン", "B-12", 1, "式", 120000, ""),
        ("会計誘導・建具CS・FRナンバー・キャッシャーサイン", "G-01/G-02/G-04", 1, "式", 180000, ""),
        ("注意喚起サイン", "G-06", 1, "式", 40000, ""),
        ("LEDサイン 取付", "G-08/G-09【支給品】", 2, "台", 35000, ""),
        ("ポスターパネル設置", "A-00g", 1, "式", 80000, ""),
        ("OPEN告知・外部サイン", "", 1, "式", 150000, "仕様要確認"),
    ]),
    ("13 電気設備工事", [
        ("テナント分電盤 電灯盤-1", "河村電器 自立 1φ3W MCB3P 225AF/200AT ELB×24", 1, "面", 980000, "盤図 U54836-01"),
        ("テナント分電盤 電灯盤-2", "河村電器 自立 1φ3W MCB3P 225AF/175AT ELB×28", 1, "面", 980000, "盤図 U54836-02"),
        ("分電盤 据付・搬入", "約106kg/面 チャンネルベース アンカー", 2, "面", 120000, ""),
        ("幹線工事", "主開閉器盤→テナント盤 CVT 2系統", 1, "式", 450000, ""),
        *[(f"照明器具 {k}", "3500K Ra85 電源ユニット込", q, "台", p, "照明器具表") for k, (q, p) in DL_BD.items()],
        *[(f"照明器具 {k}", "3500K Ra85 電源ユニット込", q, "台", p, "照明器具表") for k, (q, p) in DL_UD.items()],
        ("照明器具 BL1/BL2 直付ベースライト", "", 13, "台", 19000, "照明器具表"),
        ("照明器具 SP1 LEDスポット", "UQS22322SP", 17, "台", 33000, "照明器具表"),
        ("照明器具 SP2 LEDスポット", "ERS4815BB", 11, "台", 16000, "照明器具表"),
        ("LEDテープライト IL1", "FX50-LED 直流電源込", 1, "式", 35000, "照明器具表"),
        ("照明器具取付", "落下防止ワイヤー取付 CH3750", N_LIGHT, "台", 4500, ""),
        ("照明配線・配管", "EM-EEF・ジョイントボックス（ワゴ接続禁止）", N_LIGHT, "灯", 6500, "施工指針書 11-3"),
        ("スイッチ・3路スイッチ", "照明・空調・換気SW移設含む", 15, "箇所", 9000, "平面図"),
        ("ダブルコンセント", "配線込", 41, "個", 12000, "照明器具表"),
        ("抜け止め接地付ダブルコンセント", "配線込", 42, "個", 13500, "照明器具表"),
        ("天井裏ダブルコンセント", "", 4, "個", 9000, "照明器具表"),
        ("カットリレーコンセント 空配管・電源", "本体・信号線はB工事", 1, "箇所", 15000, "区分表"),
        ("弱電空配管（POS/TEL/FAX/LAN）", "アウトレット・プルボックス", 20, "箇所", 14000, "照明器具表"),
        ("換気扇 EF-1 電源・スイッチ", "店長室 1φ100V", 1, "式", 25000, "換気仕様書"),
        ("音響設備", "アンプ2台・天井SP10・吊下SP2・BOH SP・音量調整器・マイク", 1, "式", 980000, "AV機器配置図"),
        ("防犯（ダミーカメラ）", "SCB-4100D+取付台", 8, "台", 9000, "AV機器配置図"),
        ("試験調整・絶縁抵抗測定", "様式-20 提出", 1, "式", 80000, "施工指針書 11-3"),
    ]),
]

EXCLUSIONS = [
    ("B工事（指定業者）", "空調室内機・GHPリモコン移設、一般換気の増移設、防災設備（感知器・スピーカー・誘導灯・SP）増移設、排煙SW移設、カットリレー本体・信号線、消火器・防火区画貫通処理"),
    ("指定会社工事", "電話（NTT）、売上端末（POS）、有線放送（USEN）、個別機械警備（セコム）"),
    ("支給品", "壁面・中央什器、レジバックカウンター什器、サービスCT、有孔ボードサッカー台、内照式ロゴサイン、LEDサイン、スタンド照明ST1、金庫"),
    ("什器据付", "支給品什器の組立・据付（ご要望により別途お見積り）"),
    ("施設負担金", "総合仮設負担金、開店前準備負担金、開店時販促負担金、内装監理費 など"),
    ("その他", "設計・監理費、諸官庁申請手数料、共用部の工事、夜間特別警備費、図面にない追加・変更工事、アスベスト調査・除去"),
]

ASSUMPTIONS = [
    "数量は図面PDFのテキスト（仕上表・照明器具表・盤図・フロアマップ・建具表）から拾い出した。外周壁・間仕切の長さ、FRブース数、柱本数は図面寸法から概算しており、CAD実測で拾い直すと±10%程度変わる可能性がある。",
    "単価は2026年10月時点の首都圏商業施設テナント内装の市場単価（材工共・昼間施工）を基準にした概算。照明器具・分電盤はメーカー見積で差し替えること。",
    "施工指針書により作業は22:00〜07:00の夜間のみ（残業不可・溶接不可・搬入4t車まで）。労務の夜間割増を直接工事費の8%として計上した。",
    "工事区分は区分表（添付書類②）に従い、C工事のみを計上。A・B工事、指定会社工事、支給品は「別途工事」に記載し金額には含めない。",
    "床は既存躯体+水性ウレタン塗装の上に塩ビタイル直張り、天井（売場）はA工事PB面へAEP塗装、BOHはLGS+化粧石膏ボードの新設とした。",
    f"工程は着工を{START:%Y年%m月%d日}（夜）と仮定し、日曜夜を休みとする40夜で組んだ。実際の着工日は内装監理室の工程調整で確定する。",
]

# ---------------------------------------------------------------------------
wb = Workbook()

# ===== 明細書 =====
ws_d = wb.active
ws_d.title = "明細書"
hdr = ["No", "名称", "仕様・摘要", "数量", "単位", "単価", "金額", "数量根拠・備考", "工種"]
widths = [5, 34, 44, 9, 6, 11, 13, 40, 0.1]
ws_d["A1"] = "工事費内訳明細書"
ws_d["A1"].font = f(14, True)
ws_d["A2"] = "工事名称：GU ららぽーと富士見店（2413区画）内装C工事"
ws_d["A2"].font = f(10)
for i, (h, w) in enumerate(zip(hdr, widths), 1):
    c = ws_d.cell(row=4, column=i, value=h)
    c.font = f(10, True); c.fill = HEAD_FILL; c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_d.column_dimensions[get_column_letter(i)].width = w
ws_d.column_dimensions["I"].hidden = True
ws_d.freeze_panes = "A5"

r = 5
sub_cells = {}
for sec, rows in ITEMS:
    c = ws_d.cell(row=r, column=1, value=sec)
    ws_d.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c.font = f(10, True); c.fill = SEC_FILL
    for col in range(1, 9):
        ws_d.cell(row=r, column=col).border = BORDER
    r += 1
    first = r
    for n, (name, spec, q, u, p, note) in enumerate(rows, 1):
        vals = [n, name, spec, q, u, p, f"=ROUND(D{r}*F{r},0)", note, sec]
        for col, v in enumerate(vals, 1):
            c = ws_d.cell(row=r, column=col, value=v)
            c.font = f(9); c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=col in (2, 3, 8))
        ws_d.cell(row=r, column=4).number_format = '#,##0.0' if isinstance(q, float) and q != int(q) else '#,##0'
        ws_d.cell(row=r, column=6).number_format = YEN
        ws_d.cell(row=r, column=7).number_format = YEN
        ws_d.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws_d.cell(row=r, column=5).alignment = Alignment(horizontal="center", vertical="center")
        r += 1
    ws_d.cell(row=r, column=2, value=f"{sec[3:]} 小計").font = f(9, True)
    c = ws_d.cell(row=r, column=7, value=f"=SUM(G{first}:G{r-1})")
    c.number_format = YEN; c.font = f(9, True)
    for col in range(1, 9):
        ws_d.cell(row=r, column=col).fill = SUB_FILL
        ws_d.cell(row=r, column=col).border = BORDER
    sub_cells[sec] = f"明細書!G{r}"
    r += 2
ws_d.page_setup.orientation = "landscape"
ws_d.page_setup.paperSize = ws_d.PAPERSIZE_A4
ws_d.page_setup.fitToWidth = 1
ws_d.page_setup.fitToHeight = 0
ws_d.sheet_properties.pageSetUpPr.fitToPage = True
ws_d.print_title_rows = "4:4"

# ===== 内訳書 =====
ws_s = wb.create_sheet("内訳書", 0)
ws_s["A1"] = "工事費内訳書"
ws_s["A1"].font = f(14, True)
ws_s["A2"] = "工事名称：GU ららぽーと富士見店（2413区画）内装C工事"
for i, (h, w) in enumerate(zip(["No", "工種", "金額", "構成比", "備考"], [5, 36, 16, 9, 46]), 1):
    c = ws_s.cell(row=4, column=i, value=h)
    c.font = f(10, True); c.fill = HEAD_FILL; c.border = BORDER
    c.alignment = Alignment(horizontal="center")
    ws_s.column_dimensions[get_column_letter(i)].width = w
r = 5
ws_s.cell(row=r, column=1, value="Ⅰ 直接工事費").font = f(10, True)
r += 1
d_first = r
for sec in sub_cells:
    ws_s.cell(row=r, column=1, value=sec[:2])
    ws_s.cell(row=r, column=2, value=sec[3:])
    ws_s.cell(row=r, column=3, value=f"={sub_cells[sec]}").number_format = YEN
    r += 1
d_last = r - 1
ws_s.cell(row=r, column=2, value="直接工事費 計").font = f(10, True)
ws_s.cell(row=r, column=3, value=f"=SUM(C{d_first}:C{d_last})").number_format = YEN
ws_s.cell(row=r, column=3).font = f(10, True)
DIRECT = f"C{r}"
direct_row = r
r += 2
ws_s.cell(row=r, column=1, value="Ⅱ 共通費").font = f(10, True)
r += 1
c_first = r
rates = [("夜間作業割増", 0.08, "直接工事費×8%（22:00〜07:00作業・労務費の約25%相当）"),
         ("共通仮設費", 0.03, "直接工事費×3%"),
         ("現場管理費", 0.08, "直接工事費×8%（現場責任者 正副2名常駐・申請書類・竣工図）")]
for name, rate, note in rates:
    ws_s.cell(row=r, column=2, value=name)
    ws_s.cell(row=r, column=3, value=f"=ROUND({DIRECT}*{rate},-3)").number_format = YEN
    ws_s.cell(row=r, column=5, value=note)
    r += 1
ws_s.cell(row=r, column=2, value="一般管理費")
ws_s.cell(row=r, column=3, value=f"=ROUND(({DIRECT}+SUM(C{c_first}:C{r-1}))*0.07,-3)").number_format = YEN
ws_s.cell(row=r, column=5, value="（直接工事費＋上記共通費）×7%")
c_last = r
r += 1
ws_s.cell(row=r, column=2, value="共通費 計").font = f(10, True)
ws_s.cell(row=r, column=3, value=f"=SUM(C{c_first}:C{c_last})").number_format = YEN
ws_s.cell(row=r, column=3).font = f(10, True)
COMMON = f"C{r}"
r += 2
ws_s.cell(row=r, column=2, value="工事費 計（Ⅰ＋Ⅱ）")
ws_s.cell(row=r, column=3, value=f"={DIRECT}+{COMMON}").number_format = YEN
sub_r = r
r += 1
ws_s.cell(row=r, column=2, value="端数調整")
ws_s.cell(row=r, column=3, value=f"=-MOD(C{sub_r},10000)").number_format = YEN
adj_r = r
r += 1
ws_s.cell(row=r, column=2, value="税抜合計").font = f(11, True)
ws_s.cell(row=r, column=3, value=f"=C{sub_r}+C{adj_r}").number_format = YEN
ws_s.cell(row=r, column=3).font = f(11, True)
NET = f"内訳書!C{r}"
net_r = r
r += 1
ws_s.cell(row=r, column=2, value="消費税（10%）")
ws_s.cell(row=r, column=3, value=f"=ROUNDDOWN(C{net_r}*0.1,0)").number_format = YEN
tax_r = r
r += 1
ws_s.cell(row=r, column=2, value="税込合計").font = f(11, True)
ws_s.cell(row=r, column=3, value=f"=C{net_r}+C{tax_r}").number_format = YEN
ws_s.cell(row=r, column=3).font = f(11, True)
GROSS = f"内訳書!C{r}"
TAX = f"内訳書!C{tax_r}"
last = r
for rr in range(5, last + 1):
    for col in range(1, 6):
        ws_s.cell(row=rr, column=col).border = BORDER
        if ws_s.cell(row=rr, column=col).font.name != FONT:
            ws_s.cell(row=rr, column=col).font = f(10)
    if d_first <= rr <= d_last:
        ws_s.cell(row=rr, column=4, value=f"=C{rr}/{DIRECT}").number_format = "0.0%"
ws_s.cell(row=last + 2, column=2, value="㎡単価（税抜）")
ws_s.cell(row=last + 2, column=3, value=f"=ROUND(C{net_r}/{AREA_TOTAL},0)").number_format = YEN
ws_s.cell(row=last + 2, column=5, value=f"区画面積 {AREA_TOTAL}㎡ で割った値")
ws_s.cell(row=last + 3, column=2, value="坪単価（税抜）")
ws_s.cell(row=last + 3, column=3, value=f"=ROUND(C{net_r}/({AREA_TOTAL}/3.30579),0)").number_format = YEN
for rr in (last + 2, last + 3):
    for col in (2, 3, 5):
        ws_s.cell(row=rr, column=col).font = f(10)
ws_s.page_setup.paperSize = ws_s.PAPERSIZE_A4

# ===== 見積書（表紙）=====
ws_c = wb.create_sheet("見積書", 0)
for col, w in zip("ABCDEFGH", [3, 20, 16, 16, 16, 16, 16, 3]):
    ws_c.column_dimensions[col].width = w
ws_c.merge_cells("B2:G2")
ws_c["B2"] = "御 見 積 書"
ws_c["B2"].font = f(22, True)
ws_c["B2"].alignment = Alignment(horizontal="center")
ws_c["F4"] = "見積No."
ws_c["G4"] = "2413-001"
ws_c["F5"] = "発行日"
ws_c["G5"] = ISSUE_DATE
ws_c["G5"].number_format = "yyyy年m月d日"
ws_c.merge_cells("B6:D6")
ws_c["B6"] = "株式会社ジーユー　御中"
ws_c["B6"].font = f(14, True)
ws_c["B6"].border = Border(bottom=Side(style="medium"))
ws_c["B8"] = "下記のとおり御見積申し上げます。"
ws_c["B10"] = "御見積金額（税込）"
ws_c["B10"].font = f(12, True)
ws_c.merge_cells("C10:E10")
ws_c["C10"] = f"={GROSS}"
ws_c["C10"].number_format = '"¥"#,##0"-"'
ws_c["C10"].font = f(18, True)
ws_c["C10"].border = Border(bottom=Side(style="double"))
ws_c["B11"] = "（税抜金額）"
ws_c["C11"] = f"={NET}"
ws_c["C11"].number_format = '"¥"#,##0'
ws_c["D11"] = "消費税"
ws_c["E11"] = f"={TAX}"
ws_c["E11"].number_format = '"¥"#,##0'
info = [
    ("工事名称", "GU ららぽーと富士見店 内装工事（C工事）"),
    ("工事場所", "埼玉県富士見市山室1-1313 ららぽーと富士見 2F 2413区画"),
    ("面積", f"{AREA_TOTAL:,.1f}㎡（約{AREA_TOTAL/3.30579:.0f}坪）　CH=3,750"),
    ("工期", f"{START:%Y年%m月%d日}着工〜40夜（日曜休み）　※内装監理室の工程調整による"),
    ("支払条件", "別途協議"),
    ("見積有効期限", "発行日より30日間"),
]
r = 13
for k, v in info:
    ws_c.cell(row=r, column=2, value=k).font = f(10, True)
    ws_c.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    ws_c.cell(row=r, column=3, value=v).font = f(10)
    for col in range(2, 8):
        ws_c.cell(row=r, column=col).border = Border(bottom=thin)
    r += 1
ws_c["E20"] = "（貴社名）"
ws_c["E21"] = "（住所・電話）"
ws_c["E22"] = "担当："
for a in ("E20", "E21", "E22"):
    ws_c[a].font = f(10)
r = 24
ws_c.cell(row=r, column=2, value="【別途工事・見積除外】").font = f(10, True)
r += 1
for k, v in EXCLUSIONS:
    ws_c.cell(row=r, column=2, value="・" + k).font = f(9, True)
    ws_c.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    c = ws_c.cell(row=r, column=3, value=v)
    c.font = f(9); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws_c.row_dimensions[r].height = 30
    r += 1
r += 1
ws_c.cell(row=r, column=2, value="【見積条件】").font = f(10, True)
r += 1
for a in ASSUMPTIONS:
    ws_c.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c = ws_c.cell(row=r, column=2, value="・" + a)
    c.font = f(9); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws_c.row_dimensions[r].height = 30
    r += 1
for row in ws_c.iter_rows():
    for c in row:
        if c.font.name != FONT:
            c.font = f(10)
ws_c.page_setup.paperSize = ws_c.PAPERSIZE_A4
ws_c.sheet_properties.pageSetUpPr.fitToPage = True
ws_c.page_setup.fitToHeight = 1

# ===== 数量根拠 =====
ws_q = wb.create_sheet("数量根拠")
ws_q["A1"] = "数量拾い根拠（図面より）"
ws_q["A1"].font = f(14, True)
for i, (h, w) in enumerate(zip(["項目", "数量", "単位", "出典・算出"], [36, 10, 6, 70]), 1):
    c = ws_q.cell(row=3, column=i, value=h)
    c.font = f(10, True); c.fill = HEAD_FILL; c.border = BORDER
    ws_q.column_dimensions[get_column_letter(i)].width = w
for rr, row in enumerate(QTY_BASIS, 4):
    for col, v in enumerate(row, 1):
        c = ws_q.cell(row=rr, column=col, value=v)
        c.font = f(10); c.border = BORDER
    ws_q.cell(row=rr, column=2).number_format = '#,##0.0'
rr += 2
ws_q.cell(row=rr, column=1, value="参照図面").font = f(10, True)
for s in ["※一式.pdf（施工図一式：仕上表・平面図・建具表・照明器具表・機械仕様書・各詳細図）",
          "製作図_GUららぽーと富士見店修正盤図(210108).pdf（河村電器 電灯盤-1/-2）",
          "【L1】LPFJ_2413号室区分表.pdf（A/B/C工事区分）",
          "LPFJ_RN2021【施工指針書一式】20210108.pdf（作業時間・搬入ルール）",
          "内装設計指針書.pdf"]:
    rr += 1
    ws_q.cell(row=rr, column=1, value="・" + s).font = f(9)

# ===== 工程表 =====
def work_nights(start, n):
    d, out = start, []
    while len(out) < n:
        if d.weekday() != 6:  # 日曜夜は休み
            out.append(d)
        d += dt.timedelta(days=1)
    return out

NIGHTS = work_nights(START, 40)
# (工種, 開始夜, 終了夜, 色)
TASKS = [
    ("A・B工事引渡し確認・墨出し・養生", 1, 2, BAR_FILL),
    ("撤去（店頭サッシ・既存残置物）", 2, 5, BAR_FILL),
    ("B工事：空調・防災・排煙SW 移設（指定業者）", 3, 14, BAR2_FILL),
    ("電気：幹線・天井内配管配線・分電盤据付", 3, 14, BAR_FILL),
    ("LGS壁・間仕切下地／ファサード下地", 4, 13, BAR_FILL),
    ("BOH天井下地・防煙垂れ壁", 8, 14, BAR_FILL),
    ("ボード張り（壁・天井）", 9, 17, BAR_FILL),
    ("建具枠取付・什器取付下地", 13, 18, BAR_FILL),
    ("パテ・塗装（天井・壁 AEP）", 15, 24, BAR_FILL),
    ("クロス張り（FR）", 20, 24, BAR_FILL),
    ("照明器具・コンセント・SW 取付", 18, 28, BAR_FILL),
    ("床下地調整・塩ビタイル・巾木", 22, 29, BAR_FILL),
    ("造作家具（レジバック・FR・棚・SW造作）", 21, 30, BAR_FILL),
    ("建具吊込み・金物・見切り", 25, 31, BAR_FILL),
    ("柱巻ミラー・FRミラー", 26, 31, BAR_FILL),
    ("サイン取付（支給品含む）", 28, 34, BAR_FILL),
    ("弱電・音響・防犯（指定会社 電話/POS/有線含む）", 28, 35, BAR2_FILL),
    ("支給品什器 搬入・据付（GU様手配）", 30, 35, BAR3_FILL),
    ("試験調整・絶縁抵抗測定・照明フォーカス", 33, 36, BAR_FILL),
    ("手直し・竣工クリーニング", 35, 38, BAR_FILL),
    ("自主検査 → 内装監理室検査", 37, 38, BAR_FILL),
    ("消防検査・是正", 39, 39, BAR2_FILL),
    ("引渡し・商品搬入", 40, 40, BAR3_FILL),
]

ws_g = wb.create_sheet("施工工程表")
ws_g["A1"] = "施工工程表（C工事・夜間 22:00〜07:00）"
ws_g["A1"].font = f(14, True)
ws_g["A2"] = f"工事名称：GU ららぽーと富士見店（2413区画）　着工 {START:%Y/%m/%d}（夜）〜 {NIGHTS[-1]:%Y/%m/%d}　40夜　※日付は作業開始夜の日付、日曜夜休み"
ws_g["A2"].font = f(9)
ws_g.column_dimensions["A"].width = 4
ws_g.column_dimensions["B"].width = 44
ws_g.column_dimensions["C"].width = 6
cal = []
d = START
while d <= NIGHTS[-1]:
    cal.append(d)
    d += dt.timedelta(days=1)
night_no = {n: i + 1 for i, n in enumerate(NIGHTS)}
WD = "月火水木金土日"
for j, day in enumerate(cal):
    col = 4 + j
    ws_g.column_dimensions[get_column_letter(col)].width = 3.2
    c1 = ws_g.cell(row=4, column=col, value=f"{day.month}/{day.day}" if day.day == 1 or j == 0 or day.weekday() == 0 else "")
    c1.font = f(7)
    c2 = ws_g.cell(row=5, column=col, value=day.day)
    c3 = ws_g.cell(row=6, column=col, value=WD[day.weekday()])
    c4 = ws_g.cell(row=7, column=col, value=night_no.get(day, "休"))
    for c in (c2, c3, c4):
        c.font = f(7); c.alignment = Alignment(horizontal="center"); c.border = BORDER
        if day.weekday() == 6:
            c.fill = SUN_FILL
for rr, lab in zip((5, 6, 7), ("日", "曜", "夜")):
    ws_g.cell(row=rr, column=3, value=lab).font = f(8)
for i, h in enumerate(["No", "工種・作業内容"], 1):
    c = ws_g.cell(row=7, column=i, value=h)
    c.font = f(9, True); c.fill = HEAD_FILL; c.border = BORDER
for k, (name, s, e, fill) in enumerate(TASKS, 1):
    rr = 7 + k
    ws_g.cell(row=rr, column=1, value=k).font = f(9)
    ws_g.cell(row=rr, column=2, value=name).font = f(9)
    ws_g.cell(row=rr, column=3, value=e - s + 1).font = f(9)
    for col in (1, 2, 3):
        ws_g.cell(row=rr, column=col).border = BORDER
    sd, ed = NIGHTS[s - 1], NIGHTS[e - 1]
    for j, day in enumerate(cal):
        c = ws_g.cell(row=rr, column=4 + j)
        c.border = BORDER
        if day.weekday() == 6:
            c.fill = SUN_FILL
        elif sd <= day <= ed:
            c.fill = fill
rr = 8 + len(TASKS) + 1
for txt, fill in (("C工事（当社）", BAR_FILL), ("B工事・指定会社", BAR2_FILL), ("GU様手配・引渡し", BAR3_FILL)):
    ws_g.cell(row=rr, column=4).fill = fill
    ws_g.cell(row=rr, column=5, value=txt).font = f(8)
    rr += 1
for note in ["※ 作業は22:00入館〜07:00終了（夕礼・KY後に作業開始、残業不可）。搬入は22:00〜06:00、4t車まで、1台30分以内（北荷捌所）。",
             "※ 溶接不可。火気使用は内装監理室の承認が必要（残火確認2時間）。",
             "※ 着工10日前までに施工承認書類（様式1〜9）、7日前までに着工届・作業員名簿・A/B工事引渡し確認書を提出。"]:
    ws_g.cell(row=rr, column=2, value=note).font = f(8)
    rr += 1
ws_g.freeze_panes = "D8"
ws_g.page_setup.orientation = "landscape"
ws_g.page_setup.paperSize = ws_g.PAPERSIZE_A3
ws_g.sheet_properties.pageSetUpPr.fitToPage = True
ws_g.page_setup.fitToHeight = 1

# ===== 全体工程表（週単位）=====
ws_w = wb.create_sheet("全体工程表")
ws_w["A1"] = "全体工程表（見積〜引渡し・週単位）"
ws_w["A1"].font = f(14, True)
ws_w.column_dimensions["A"].width = 4
ws_w.column_dimensions["B"].width = 44
W0 = dt.date(2026, 10, 19)  # 月曜
weeks = [W0 + dt.timedelta(weeks=i) for i in range(20)]
for j, w in enumerate(weeks):
    col = 3 + j
    ws_w.column_dimensions[get_column_letter(col)].width = 6
    c = ws_w.cell(row=3, column=col, value=f"{w.month}/{w.day}")
    c.font = f(8); c.alignment = Alignment(horizontal="center"); c.border = BORDER; c.fill = HEAD_FILL
for i, h in enumerate(["No", "項目"], 1):
    c = ws_w.cell(row=3, column=i, value=h)
    c.font = f(9, True); c.fill = HEAD_FILL; c.border = BORDER
D = dt.date
PHASES = [
    ("図面受領・積算・見積提出", D(2026, 10, 21), D(2026, 10, 30), BAR_FILL),
    ("見積調整（VE・数量精査）・契約", D(2026, 11, 2), D(2026, 11, 13), BAR_FILL),
    ("実施設計図 内装監理室承認・製作図作成/承認", D(2026, 11, 9), D(2026, 11, 27), BAR_FILL),
    ("照明器具・分電盤 発注〜納品（盤製作 約5週）", D(2026, 11, 16), D(2026, 12, 25), BAR_FILL),
    ("建具・造作家具・ミラー・サイン 工場製作", D(2026, 11, 30), D(2027, 1, 8), BAR_FILL),
    ("施工承認申請（様式1〜9：着工10日前）", D(2026, 12, 21), D(2026, 12, 25), BAR2_FILL),
    ("着工届・作業員名簿・A/B工事引渡し確認（7日前）", D(2027, 1, 5), D(2027, 1, 8), BAR2_FILL),
    ("C工事 施工（40夜）", NIGHTS[0], NIGHTS[-1], BAR_FILL),
    ("内装監理室検査・消防検査", NIGHTS[36], NIGHTS[38], BAR2_FILL),
    ("引渡し・商品搬入・開店準備（GU様）", NIGHTS[-1], NIGHTS[-1] + dt.timedelta(days=6), BAR3_FILL),
    ("竣工図・竣工写真・シックハウス書類提出（竣工後1ヶ月以内）", NIGHTS[-1], NIGHTS[-1] + dt.timedelta(days=27), BAR2_FILL),
]
for k, (name, s, e, fill) in enumerate(PHASES, 1):
    rr = 3 + k
    ws_w.cell(row=rr, column=1, value=k).font = f(9)
    ws_w.cell(row=rr, column=2, value=name).font = f(9)
    for col in (1, 2):
        ws_w.cell(row=rr, column=col).border = BORDER
    for j, w in enumerate(weeks):
        c = ws_w.cell(row=rr, column=3 + j)
        c.border = BORDER
        if s <= w + dt.timedelta(days=6) and e >= w:
            c.fill = fill
rr = 4 + len(PHASES) + 1
ws_w.cell(row=rr, column=2, value="※ 年末年始（12/28〜1/4）は館の工事休止を想定し、着工を1/12とした。").font = f(8)
ws_w.page_setup.orientation = "landscape"
ws_w.page_setup.paperSize = ws_w.PAPERSIZE_A4
ws_w.sheet_properties.pageSetUpPr.fitToPage = True
ws_w.page_setup.fitToHeight = 1

wb.save(OUT)
print("saved", OUT)
