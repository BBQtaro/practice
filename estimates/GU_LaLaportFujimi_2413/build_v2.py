# -*- coding: utf-8 -*-
"""GU ららぽーと富士見店（2413区画）C工事 見積 v2（実物件見積 2021/01 との答え合わせ反映版）

python3 build_v2.py  ->  見積書・工程表・拾い出し根拠 の3ファイル
v1 からの主な変更:
  - 照明器具・床塩ビタイル材・カウンター・LEDサインは GU 支給品として施工費のみ計上
  - 大型造作（レジバック・FR・ストック棚等）を削除し、GU 標準化項目で計上
  - BOH 天井下地・化粧ボードを削除（GB-D は BOH 壁仕上げ。天井は全面 AEP）
  - コーナーガード・BOH クロス・付帯（指定産廃・揚重・駐車場）・夜間警備を追加
  - 単価は 2021 年実績売価 × 物価補正係数 K（2026 年水準）
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

K = 1.25                      # 2021→2026 労務・資材の上昇率（仮定）
ISSUE_DATE = dt.date(2026, 10, 21)
START = dt.date(2027, 1, 12)  # C工事 着工（初夜）仮定
AREA = 1064.0
FONT = "游ゴシック"
PREFIX = "GU富士見_2413_"

thin = Side(style="thin", color="808080")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD = PatternFill("solid", fgColor="DCE6F1")
SEC = PatternFill("solid", fgColor="F2F2F2")
SUB = PatternFill("solid", fgColor="FFF2CC")
BAR = PatternFill("solid", fgColor="4F81BD")
BAR2 = PatternFill("solid", fgColor="F79646")
BAR3 = PatternFill("solid", fgColor="9BBB59")
SUN = PatternFill("solid", fgColor="FDE9D9")
YEN = "#,##0"


def f(size=10, bold=False):
    return Font(name=FONT, size=size, bold=bold)


def price(base):
    v = base * K
    return int(round(v, -1 if v < 10000 else -2))


# 根拠の略号
D = "図面"        # 2413区画 施工図一式・照明器具表・盤図から拾い
A = "実績"        # 2021/01 実物件見積の数量
DA = "図面・実績"  # 図面と実績でほぼ一致

# (工種, [(名称, 仕様, 数量, 単位, 2021基準単価, 根拠, 拾い出しメモ)])
ITEMS = [
    ("01 仮設工事", [
        ("墨出し", "事前床レベル墨出し含む", AREA, "㎡", 100, D, "照明器具表 床面積1,064㎡"),
        ("現場養生", "", AREA, "㎡", 100, D, "同上"),
        ("清掃片付け", "日々", AREA, "㎡", 100, D, "同上"),
        ("場内小運搬", "", AREA, "㎡", 100, D, "同上"),
        ("引渡し清掃", "", AREA, "㎡", 300, D, "同上"),
        ("発生残材処分", "", 40, "m3", 8400, A, ""),
        ("搬出入路・ELV養生", "日々掛け払い 100m×2人", 12, "回", 15000, A, ""),
        ("ロール養生材", "100m巻", 12, "本", 5000, A, ""),
        ("搬出入床コンパネ養生", "160m", 1, "式", 250000, A, "施工指針書：300kg以上は9mmポリプレート＋コンパネ"),
        ("仮設分電盤・仮設照明・仮設資材", "漏電遮断器付 電源送り含む", 1, "式", 36000, A, "施工指針書：仮設照明の準備なし"),
        ("アップスター", "", 2, "台", 29000, A, ""),
        ("フリーダムステージ", "CH3750 作業床", 10, "台", 40000, A, ""),
        ("立馬", "", 5, "台", 15000, A, ""),
        ("ワークスペースライト", "", 6, "台", 20000, A, ""),
        ("足場運搬費", "", 3, "台", 45000, A, ""),
        ("仮囲い撤去・清掃", "", 110, "㎡", 1600, A, "ファサード20.8m×H≒5m"),
    ]),
    ("02 軽鉄・ボード工事", [
        ("壁面 LGS下地 ST100", "@303", 691, "㎡", 2000, A, "外周C工事壁（図面概算 110m×3.75m=413㎡＋FR・レジバック）"),
        ("壁面 LGS下地 ST65", "@303 BOH間仕切", 285, "㎡", 1830, A, "図面概算 75m×3.75m=281㎡"),
        ("柱 LGS下地 ST100", "柱巻き", 156, "㎡", 1830, A, ""),
        ("壁面 PB12.5 張り", "不燃 レジバック・FR壁面3×8", 965, "㎡", 1240, A, "実績1,238㎡−BOH化粧ボード変更分273㎡"),
        ("柱 PB12.5 張り", "不燃", 156, "㎡", 1240, A, ""),
        ("ファサードゲート LGS下地", "", 66, "㎡", 2500, DA, "図面 ボーダー重量計算 PB計39.5㎡＋見付"),
        ("ファサードゲート PB張り", "不燃12.5", 66, "㎡", 1500, DA, ""),
        ("ファサード補強", "吊りボルト@900", 27, "m", 3500, A, ""),
        ("天井点検口 売場", "450角 アルミフレーム", 16, "ヶ所", 4550, A, ""),
        ("天井点検口 BOH", "450角 アルミフレーム", 8, "ヶ所", 4550, A, ""),
        ("サイン用下地", "", 5, "ヶ所", 3500, A, ""),
        ("建具開口補強", "", 12, "ヶ所", 4000, DA, "建具表"),
        ("DL開口補強", "半数として", 198, "ヶ所", 1000, D, "照明器具表 DL396台×1/2"),
        ("壁面排煙開口補強", "", 19, "ヶ所", 5000, A, ""),
        ("防煙垂壁下地追加", "テンションタイプ", 10, "箇所", 15000, A, "平面図 防煙垂れ壁H=500"),
        ("各所補足スタッド・FR手摺下地", "", 22, "ヶ所", 1200, A, ""),
        ("補正室 天井ケイカル増し張り", "6mm＋コーナーフクビ", 1, "式", 25000, A, ""),
        ("L型コーナー", "塩ビコーナー 材工", 275, "m", 400, A, ""),
        ("ピクチャーレール下地補強", "", 1, "式", 40000, D, "B-09a"),
        ("壁面・柱埋込スリット", "ASF-52 L2400＋専用下地PST-5＋取付", 8, "本", 8000, A, ""),
        ("施工副資材・運搬搬入", "", 1, "式", 195000, A, ""),
        ("廃材処理", "", 20, "m3", 8400, A, ""),
    ]),
    ("03 床タイル工事", [
        ("売場 塩ビタイル施工", "ES3001-45 馬貼り【材料GU支給】", 876, "㎡", 1050, DA, "図面 売場855㎡＋FRブース外"),
        ("FRブース内 Pタイル施工", "WBH-603【材料GU支給】", 12, "㎡", 1300, DA, "FR小8＋大1＋HC1"),
        ("BOH 塩ビタイル施工", "P-41【材料GU支給】", 169, "㎡", 1050, DA, "図面 BOH5室173㎡（SW含め要確認）"),
        ("床下地処理", "アースシール程度", 1, "式", 30000, A, ""),
        ("専用接着剤", "アクリル系", 1057, "㎡", 300, DA, ""),
        ("各所床見切り", "", 1, "式", 45000, A, ""),
        ("運搬搬入", "", 1, "式", 110000, A, ""),
    ]),
    ("04 塗装工事", [
        ("天井 AEP", "N-95ツヤ消し 寒冷紗パテ含む（BOH含む全面）", 1065, "㎡", 1600, DA, "区画面積"),
        ("壁面 AEP", "N-95 3分ツヤ 寒冷紗パテ含む", 645, "㎡", 1600, A, ""),
        ("柱面 AEP", "N-95 3分ツヤ", 64, "㎡", 1600, A, ""),
        ("ファサードゲート AEP", "N-95 3分ツヤ", 59, "㎡", 1600, A, ""),
        ("A工事避難扉 SOP", "N-87 原状復旧対象", 3, "枚", 15000, D, "平面図 注記"),
        ("シャッターガイド SOP", "", 3, "人工", 30000, A, ""),
        ("塗装養生", "", 1065, "㎡", 50, D, ""),
        ("雑塗装・補修・副資材・処分", "", 1, "式", 97000, A, ""),
    ]),
    ("05 内装仕上げ工事", [
        ("売場・FR ソフト巾木", "HB-89 H100", 182, "m", 300, A, ""),
        ("売場 塩ビFB", "", 65, "m", 280, A, ""),
        ("FR・補正室 クロス（白）", "LW-4630（補正室はボード素地）", 131, "㎡", 1250, A, "実績194㎡−補正室素地63㎡"),
        ("スタッフルーム・店長室 化粧ボード（木目）", "材工", 273, "㎡", 1900, A, "仕上表 GB-D"),
        ("BOH ソフト巾木", "TH75-96 H75", 145, "m", 300, A, ""),
        ("倉庫・備品庫 巾木", "WTH300-96 H300", 75, "m", 900, A, ""),
        ("FRカーテン W1000", "サンゲツ CK-125 H2100", 8, "枚", 13000, A, ""),
        ("FRカーテン W1500（大型）", "サンゲツ CK-125 H2100", 1, "枚", 15000, A, ""),
        ("カーテンレール・タッセル", "TOSO", 9, "ヶ所", 2800, A, ""),
        ("避難通路ドットシール", "φ30 600P＋取付", 286, "m", 230, A, ""),
        ("床用イエローテープ", "3M PM1981 材工", 1, "式", 15000, D, "仕上表 分電盤まわり"),
        ("ストック アルミシート貼り 壁", "材料別途", 165, "㎡", 1200, A, "B-08"),
        ("ストック アルミシート貼り 天井", "材料支給", 38, "㎡", 4400, A, "接着材・テープ含む"),
        ("SW ロールスクリーン W2600×H4000", "手動 化粧BOX共", 6, "台", 104000, A, "B-11（実績は追加工事）"),
        ("SW ロールスクリーン W1800×H4000", "手動 化粧BOX共", 6, "台", 72000, A, ""),
        ("ロールスクリーン取付・運搬", "", 1, "式", 345000, A, ""),
    ]),
    ("06 金物工事", [
        ("ファサードシャッター塞ぎパネル", "スチールメラ焼 白", 9, "㎡", 32000, A, ""),
        ("同上 取付", "", 2, "人工", 30000, A, ""),
        ("壁面コーナーガード", "Stメラ焼 16角 H4000", 67, "本", 11000, A, ""),
        ("壁面ミラー用 メタカラー巾木", "H150 L2700", 8, "本", 7200, A, ""),
        ("ピクチャーレール", "白 L3000", 5, "本", 7200, DA, "B-09a"),
        ("CAフック・ミニハンガー・ストップ・Sカン他", "", 1, "式", 79000, A, ""),
        ("FRミラーエッジ・靴べら・コートフック", "", 1, "式", 36000, A, ""),
        ("タペストリーBOX（壁面埋込）", "W1020×H2230 取付・ケイカル下地・運搬込", 2, "箇所", 210000, DA, "B-14"),
        ("金庫固定金物", "スチール サビ止め", 1, "組", 9000, D, "C-05"),
        ("搬入・運搬・副資材", "", 1, "式", 75000, A, ""),
    ]),
    ("07 木工造作工事", [
        ("レジバック下地コンパネ張り", "", 1, "式", 50000, A, "C-03（造作本体は支給什器）"),
        ("HC用手摺", "TOTO", 2, "組", 22000, A, ""),
        ("POPケース用 木工下地", "", 21, "ヶ所", 3000, A, ""),
        ("マスターキー対応", "", 1, "式", 30000, D, "建具表 注記"),
        ("補正室ミシン台上部棚板", "W1200×D350 ブラケット・下地共", 3, "台", 25000, D, "F-03"),
        ("清掃用具掛・5ポイントフック・キーボックス", "", 1, "式", 36000, A, ""),
        ("現場取付・副資材・運搬", "", 1, "式", 225000, A, ""),
    ]),
    ("08 建具工事", [
        ("木製建具 片開き", "アイカ ポリメラ 三方枠共", 4, "組", 112000, DA, "建具表 WD-1A〜1D"),
        ("木製建具 親親開き", "召合せ フラットバー共", 1, "組", 181000, DA, "建具表 WD-3A"),
        ("金庫隠し扉・天板", "", 1, "組", 55000, DA, "建具表 WD-5A"),
        ("金庫室壁面造作", "金庫下台共", 1, "式", 20000, A, ""),
        ("スチール片開き 鍵付 枠共", "W950×H2125 クローザー・戸当・本締錠・押棒", 1, "組", 140000, DA, "建具表 SD-1A"),
        ("FR小ブース三方枠（カーテン）", "オーク無垢＋素地CL W750", 8, "組", 60000, DA, "GU標準 FR小8"),
        ("FRブース 建具付三方枠", "ポリメラ", 1, "組", 128000, DA, "建具表 WD-4A"),
        ("FR HCブース三方枠（カーテン）", "オーク無垢 W1450", 1, "ヶ所", 68000, DA, ""),
        ("建具枠保護ガード", "", 1, "式", 5000, A, ""),
        ("現場取付・運搬", "", 1, "式", 310000, A, ""),
    ]),
    ("09 ガラス工事", [
        ("柱巻きミラー", "W600×H1800 t5 四方ミラーエッジ", 14, "枚", 9200, A, ""),
        ("壁面出柱ミラー", "W600×H1800 t5", 5, "枚", 9200, A, ""),
        ("ミラーエッジ", "25×25 Lアングル黒皮風", 19, "ヶ所", 14000, A, ""),
        ("ミラー持ち出しPB", "施工費込", 19, "箇所", 4300, A, ""),
        ("FRミラー", "t5 W700×H1800", 10, "枚", 9200, DA, "FR10ブース"),
        ("スタッフルームミラー", "t5 W600×H1800", 1, "枚", 9200, A, ""),
        ("雑シリコンコーキング", "", 300, "m", 350, A, ""),
        ("ガラス施工費", "", 14, "人工", 30000, A, ""),
        ("運搬・副資材", "", 1, "式", 35000, A, ""),
    ]),
    ("10 サイン工事", [
        ("店頭LEDサイン 下地出し・取付", "GU LED 600□【支給品】", 5, "台", 16500, DA, "G-10"),
        ("店頭LEDサイン 全埋込", "450□【支給品】下地・枠金物・エースライト・取付", 2, "台", 115000, A, ""),
        ("サッシ内LEDサイン 取付", "1200□【支給品】", 2, "人工", 30000, A, "G-05"),
        ("LEDサイン用スチールボックス", "600×600 設置費込 開口補強共", 1, "ヶ所", 100000, A, ""),
        ("レジバックLED・キャッシャーLEDサイン 取付", "【支給品】化粧パイプ・下地補強・取付", 1, "式", 290000, DA, "G-08 / G-09"),
        ("通路カテゴリー誘導サイン", "木工＋アクリル t30", 6, "組", 120000, A, "GU標準化項目"),
        ("FR入口ボーダーサイン", "木工＋アクリル", 1, "組", 80000, A, ""),
        ("柱面CSサイン", "CS切り文字", 10, "組", 10000, A, ""),
        ("FRナンバープレート・STAFF ONLY・ミラードット", "", 1, "式", 25000, DA, "G-02"),
        ("補正価格ケース", "日本コバック", 20, "個", 2740, A, ""),
        ("サイン取付費", "", 9, "人工", 30000, A, ""),
        ("運搬・副資材", "", 1, "式", 120000, A, ""),
    ]),
    ("11 照明・電気工事", [
        ("電灯分電盤", "河村電器 自立 マグネット盤・アースバー（メーター別途）", 2, "面", 500000, DA, "盤図 電灯盤-1/-2"),
        ("分電盤 幹線接続・取付下地補強", "", 1, "式", 55000, A, ""),
        ("ダウンライト 開口・取付・配線", "【器具GU支給】", 396, "ヶ所", 4500, D, "照明器具表 BD262＋UD134（実績375）"),
        ("ダウンライト 落下防止ワイヤー", "ワイヤー支給", 396, "ヶ所", 700, D, ""),
        ("直付器具・スポット 取付・配線", "BL13＋SP28【器具GU支給】", 41, "ヶ所", 4000, D, "照明器具表"),
        ("間接照明 取付・配線・設置金物", "壁面スリット部", 22, "ヶ所", 6000, A, ""),
        ("配線ダクト・レースウェイ", "器具取付・付属品・支持材共", 1, "式", 231000, A, ""),
        ("スイッチ 取付・配線（3路含む）", "カバー共", 36, "ヶ所", 4500, A, "平面図 SW位置"),
        ("コンセント 取付・配線", "防犯ゲート・RFID含む", 87, "ヶ所", 4500, D, "照明器具表 41＋42＋4"),
        ("カットリレーコンセント", "本体＋配線（信号線つなぎ別途）", 1, "ヶ所", 80000, DA, ""),
        ("セルフレジ埋設コンセント", "", 1, "式", 25000, A, ""),
        ("LEDサイン用 電源送り", "", 5, "ヶ所", 12000, A, ""),
        ("分岐回路工事", "回路毎の配線 材工", 49, "回路", 5500, D, "盤図 分岐回路数"),
        ("分岐回路工事 空調・換気・動力", "", 9, "回路", 6500, A, ""),
        ("弱電空配管", "NTT・有線・POS・LAN・音響 22〜28Φ", 51, "ヶ所", 6000, DA, "照明器具表 POS/TEL/FAX 20＋各所"),
        ("同上用 天井点検口", "450角", 3, "ヶ所", 5500, A, ""),
        ("フロント区画貫通処理", "フィブロック", 1, "式", 100000, A, ""),
        ("雑材・副資材・運搬・梱包材処理", "", 1, "式", 250000, A, ""),
    ]),
    ("12 標準化項目（GU標準仕様）", [
        ("身障者ブース スツール", "張地オクタビア 台座RK-6306 取付・運搬込", 1, "台", 410000, A, ""),
        ("壁面持ち出しスリット造作", "ロイヤルスタッド＋PB・補強・AEP N-87", 40, "㎡", 13000, A, ""),
        ("ロイヤルスリット＠300", "APS-14 N-87特注塗装 L2000 下地・取付込", 60, "本", 12500, A, ""),
        ("小口押さえ金物", "上下・両側面曲げ加工", 1, "式", 182000, A, ""),
        ("ベース壁面クロス", "ツヤ消し スリット以外", 55, "㎡", 1350, A, ""),
        ("壁面木棚", "オーク D450×W900 取付共", 8, "枚", 16800, A, ""),
        ("壁面ヘリンボーンパネル", "タモ柾目 ウレタンCL 加工・見切り・取付共", 8, "枚", 35000, A, ""),
        ("壁面ピクトサイン", "取付共", 1, "ヶ所", 60000, A, ""),
        ("支給カウンター類 据付", "レジ・サービス・バックカウンター・PCカウンター", 3, "人工", 30000, A, "Y-01/Y-02"),
        ("運搬・諸経費", "", 1, "式", 300000, A, ""),
    ]),
    ("13 夜間工事費", [
        ("夜間工事割増", "22:00〜07:00 全工種", 444, "人工", 7500, A, "実績 工種別夜間人工の合計"),
        ("夜間警備費", "", 36, "日", 30000, A, ""),
    ]),
    ("14 付帯工事", [
        ("現場指定産廃処理費", "", 1, "式", 400000, A, ""),
        ("特殊揚重費", "", 1, "式", 300000, A, ""),
        ("作業員駐車場費", "P8・P5 月極 20台×3ヶ月", 1, "式", 300000, A, "施工指針書"),
    ]),
    ("15 諸経費", [
        ("現場管理費", "正副2名常駐", 1, "式", 2592000, A, ""),
        ("現場監理・調整業務", "", 1, "式", 600000, A, ""),
        ("現場諸経費", "", 1, "式", 270000, A, ""),
        ("施工承認図作成費", "", 1, "式", 100000, A, ""),
        ("竣工図書作成費", "竣工写真含む", 1, "式", 100000, A, "施工指針書 竣工後1ヶ月以内"),
        ("会社経費", "", 1, "式", 650000, A, ""),
        ("1年点検調査費", "", 1, "式", 100000, A, ""),
    ]),
]

EXCLUSIONS = [
    ("GU支給品", "照明器具・ランプ、床塩ビタイル材、アルミシート材、レジ・サービス・バックカウンター、壁面・中央什器、LEDサイン類、金庫、スタンド照明"),
    ("B工事・指定会社", "空調・換気・防災設備の増移設、排煙SW移設、カットリレー信号線、電話（NTT）、POS、有線（USEN）、機械警備"),
    ("その他", "施設負担金、設計・監理費、サッシ面遮光シート、店頭サッシ撤去（工事区分 要確認）、図面にない追加・変更工事"),
]

NOTES = [
    "単価は2021年1月の同物件実績（売価）に物価補正係数 K=%.2f を掛けた2026年水準の概算です。" % K,
    "数量は図面（照明器具表・盤図・建具表・フロアマップ）と同物件の実績数量を併用しました。根拠は拾い出し根拠資料を参照。",
    "作業は館ルールにより22:00〜07:00の夜間のみ（残業不可・溶接不可・搬入4t車まで）。",
]


# ---------------------------------------------------------------------------
def head_row(ws, row, heads, widths):
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = f(10, True); c.fill = HEAD; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w


def a4(ws, landscape=False, paper=None):
    ws.page_setup.paperSize = paper or ws.PAPERSIZE_A4
    if landscape:
        ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


# ===== 1. 見積書 =====
wb = Workbook()
ws_d = wb.active
ws_d.title = "明細書"
ws_d["A1"] = "工事費内訳明細書"
ws_d["A1"].font = f(14, True)
ws_d["A2"] = "GU ららぽーと富士見店（2413区画）内装C工事"
head_row(ws_d, 4, ["No", "名称", "仕様・摘要", "数量", "単位", "単価", "金額", "備考"], [5, 34, 42, 9, 6, 11, 13, 30])
ws_d.freeze_panes = "A5"
r = 5
subs = []
for sec, rows in ITEMS:
    ws_d.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c = ws_d.cell(row=r, column=1, value=sec)
    c.font = f(10, True); c.fill = SEC
    r += 1
    first = r
    for n, (name, spec, q, u, base, src, memo) in enumerate(rows, 1):
        vals = [n, name, spec, q, u, price(base), f"=ROUND(D{r}*F{r},0)", memo]
        for col, v in enumerate(vals, 1):
            c = ws_d.cell(row=r, column=col, value=v)
            c.font = f(9); c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=col in (2, 3, 8),
                                    horizontal="center" if col in (1, 5) else None)
        for col in (4, 6, 7):
            ws_d.cell(row=r, column=col).number_format = YEN
        r += 1
    ws_d.cell(row=r, column=2, value=f"{sec[3:]} 小計").font = f(9, True)
    c = ws_d.cell(row=r, column=7, value=f"=SUM(G{first}:G{r-1})")
    c.number_format = YEN; c.font = f(9, True)
    for col in range(1, 9):
        ws_d.cell(row=r, column=col).fill = SUB
        ws_d.cell(row=r, column=col).border = BORDER
    subs.append((sec, f"明細書!G{r}"))
    r += 2
a4(ws_d, True)
ws_d.print_title_rows = "4:4"

ws_s = wb.create_sheet("内訳書", 0)
ws_s["A1"] = "工事費内訳書"
ws_s["A1"].font = f(14, True)
ws_s["A2"] = "GU ららぽーと富士見店（2413区画）内装C工事"
head_row(ws_s, 4, ["No", "工種", "金額", "構成比"], [5, 36, 16, 9])
r = 5
for sec, ref in subs:
    ws_s.cell(row=r, column=1, value=sec[:2])
    ws_s.cell(row=r, column=2, value=sec[3:])
    ws_s.cell(row=r, column=3, value=f"={ref}").number_format = YEN
    ws_s.cell(row=r, column=4, value=f"=C{r}/C{5+len(subs)}").number_format = "0.0%"
    r += 1
rows_fin = [("小計", f"=SUM(C5:C{r-1})"), ("端数調整", f"=-MOD(C{r},10000)"),
            ("税抜合計", f"=C{r}+C{r+1}"), ("消費税（10%）", f"=ROUNDDOWN(C{r+2}*0.1,0)"),
            ("税込合計", f"=C{r+2}+C{r+3}")]
for i, (lab, fm) in enumerate(rows_fin):
    ws_s.cell(row=r + i, column=2, value=lab).font = f(10, lab.endswith("合計"))
    c = ws_s.cell(row=r + i, column=3, value=fm)
    c.number_format = YEN; c.font = f(10, lab.endswith("合計"))
NET, TAX, GROSS = f"内訳書!C{r+2}", f"内訳書!C{r+3}", f"内訳書!C{r+4}"
last = r + 4
for rr in range(5, last + 1):
    for col in range(1, 5):
        cell = ws_s.cell(row=rr, column=col)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = f(10)
ws_s.cell(row=last + 2, column=2, value="㎡単価（税抜・1,064㎡）").font = f(10)
ws_s.cell(row=last + 2, column=3, value=f"=ROUND(C{r+2}/{AREA},0)").number_format = YEN
ws_s.cell(row=last + 3, column=2, value="坪単価（税抜）").font = f(10)
ws_s.cell(row=last + 3, column=3, value=f"=ROUND(C{r+2}/({AREA}/3.30579),0)").number_format = YEN
ws_s.cell(row=last + 4, column=2, value="参考：2021年実績（税抜）").font = f(10)
ws_s.cell(row=last + 4, column=3, value=42760000).number_format = YEN
a4(ws_s)

ws_c = wb.create_sheet("見積書", 0)
for col, w in zip("ABCDEFGH", [3, 20, 16, 16, 16, 16, 16, 3]):
    ws_c.column_dimensions[col].width = w
ws_c.merge_cells("B2:G2")
ws_c["B2"] = "御 見 積 書"
ws_c["B2"].font = f(22, True)
ws_c["B2"].alignment = Alignment(horizontal="center")
ws_c["F4"], ws_c["G4"] = "見積No.", "2413-002"
ws_c["F5"], ws_c["G5"] = "発行日", ISSUE_DATE
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
ws_c["B11"], ws_c["C11"], ws_c["D11"], ws_c["E11"] = "（税抜金額）", f"={NET}", "消費税", f"={TAX}"
ws_c["C11"].number_format = ws_c["E11"].number_format = '"¥"#,##0'
info = [("工事名称", "GU ららぽーと富士見店 新装工事（C工事）"),
        ("工事場所", "埼玉県富士見市山室1-1313 ららぽーと富士見 2F 2413区画"),
        ("面積", f"{AREA:,.1f}㎡（約{AREA/3.30579:.0f}坪）　CH=3,750"),
        ("工期", f"{START:%Y年%m月%d日}着工〜40夜（夜間工事）※内装監理室の工程調整による"),
        ("支払条件", "完了時一括（別途協議）"),
        ("見積有効期限", "発行日より3ヶ月")]
r = 13
for k, v in info:
    ws_c.cell(row=r, column=2, value=k).font = f(10, True)
    ws_c.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    ws_c.cell(row=r, column=3, value=v)
    for col in range(2, 8):
        ws_c.cell(row=r, column=col).border = Border(bottom=thin)
    r += 1
ws_c["E20"], ws_c["E21"], ws_c["E22"] = "株式会社スペース", "（住所・電話）", "担当："
r = 24
ws_c.cell(row=r, column=2, value="【別途工事・見積除外】").font = f(10, True)
for k, v in EXCLUSIONS:
    r += 1
    ws_c.cell(row=r, column=2, value="・" + k).font = f(9, True)
    ws_c.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    c = ws_c.cell(row=r, column=3, value=v)
    c.font = f(9); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws_c.row_dimensions[r].height = 30
r += 2
ws_c.cell(row=r, column=2, value="【見積条件】").font = f(10, True)
for s in NOTES:
    r += 1
    ws_c.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c = ws_c.cell(row=r, column=2, value="・" + s)
    c.font = f(9); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws_c.row_dimensions[r].height = 26
for row in ws_c.iter_rows():
    for c in row:
        if c.font.name != FONT:
            c.font = f(10)
a4(ws_c)
wb.save(PREFIX + "見積書_v2.xlsx")


# ===== 2. 工程表 =====
def work_nights(start, n):
    d, out = start, []
    while len(out) < n:
        if d.weekday() != 6:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


NIGHTS = work_nights(START, 40)
TASKS = [
    ("A・B工事引渡し確認・墨出し・養生・足場搬入", 1, 2, BAR),
    ("B工事：空調・防災・排煙SW 移設（指定業者）", 3, 14, BAR2),
    ("電気：分電盤据付・幹線接続・天井内配線・弱電空配管", 3, 14, BAR),
    ("LGS壁・柱・間仕切下地／ファサードゲート下地", 3, 12, BAR),
    ("PB張り（壁・柱・ファサード）・DL開口補強", 8, 16, BAR),
    ("スリット造作・タペストリーBOX・建具枠取付", 13, 19, BAR),
    ("パテ・塗装（天井全面・壁・柱 AEP）", 15, 25, BAR),
    ("化粧ボード・クロス・アルミシート", 20, 25, BAR),
    ("照明器具（支給）取付・コンセント・SW", 19, 29, BAR),
    ("床下地処理・塩ビタイル（材料支給）・巾木", 23, 30, BAR),
    ("建具吊込み・金物・コーナーガード・見切り", 25, 31, BAR),
    ("ミラー・FR三方枠・カーテン・ロールスクリーン", 26, 32, BAR),
    ("サイン取付（支給LEDサイン含む）", 28, 34, BAR),
    ("弱電・指定会社（電話/POS/有線/警備）", 29, 35, BAR2),
    ("支給什器・カウンター 搬入据付（GU様手配）", 31, 36, BAR3),
    ("試験調整・絶縁抵抗測定・ライティング調整", 34, 36, BAR),
    ("手直し・引渡し清掃", 36, 38, BAR),
    ("自主検査 → 内装監理室検査・消防検査", 37, 39, BAR2),
    ("引渡し・商品搬入", 40, 40, BAR3),
]
wb = Workbook()
ws_g = wb.active
ws_g.title = "施工工程表"
ws_g["A1"] = "施工工程表（C工事・夜間 22:00〜07:00）"
ws_g["A1"].font = f(14, True)
ws_g["A2"] = f"GU ららぽーと富士見店（2413区画）　{START:%Y/%m/%d}（夜）着工〜{NIGHTS[-1]:%Y/%m/%d}　40夜（日曜夜休み）"
ws_g["A2"].font = f(9)
ws_g.column_dimensions["A"].width = 4
ws_g.column_dimensions["B"].width = 46
ws_g.column_dimensions["C"].width = 5
cal = [START + dt.timedelta(days=i) for i in range((NIGHTS[-1] - START).days + 1)]
nno = {n: i + 1 for i, n in enumerate(NIGHTS)}
WD = "月火水木金土日"
for j, day in enumerate(cal):
    col = 4 + j
    ws_g.column_dimensions[get_column_letter(col)].width = 3.2
    ws_g.cell(row=4, column=col, value=f"{day.month}/{day.day}" if j == 0 or day.weekday() == 0 else None).font = f(7)
    for rr, v in ((5, day.day), (6, WD[day.weekday()]), (7, nno.get(day, "休"))):
        c = ws_g.cell(row=rr, column=col, value=v)
        c.font = f(7); c.alignment = Alignment(horizontal="center"); c.border = BORDER
        if day.weekday() == 6:
            c.fill = SUN
for rr, lab in ((5, "日"), (6, "曜"), (7, "夜")):
    ws_g.cell(row=rr, column=3, value=lab).font = f(8)
for i, h in enumerate(["No", "工種・作業内容"], 1):
    c = ws_g.cell(row=7, column=i, value=h)
    c.font = f(9, True); c.fill = HEAD; c.border = BORDER
for k, (name, s, e, fill) in enumerate(TASKS, 1):
    rr = 7 + k
    for col, v in ((1, k), (2, name), (3, e - s + 1)):
        c = ws_g.cell(row=rr, column=col, value=v)
        c.font = f(9); c.border = BORDER
    for j, day in enumerate(cal):
        c = ws_g.cell(row=rr, column=4 + j)
        c.border = BORDER
        if day.weekday() == 6:
            c.fill = SUN
        elif NIGHTS[s - 1] <= day <= NIGHTS[e - 1]:
            c.fill = fill
rr = 9 + len(TASKS)
for txt, fill in (("C工事（当社）", BAR), ("B工事・指定会社・検査", BAR2), ("GU様手配・引渡し", BAR3)):
    ws_g.cell(row=rr, column=4).fill = fill
    ws_g.cell(row=rr, column=5, value=txt).font = f(8)
    rr += 1
for note in ["※ 22:00入館〜07:00終了（夕礼・KY後に作業開始、残業不可）。搬入は北荷捌所 22:00〜06:00、4t車まで、1台30分以内。",
             "※ 溶接不可。火気は内装監理室承認（残火確認2時間）。夜間警備は実績同様36日を計上。",
             "※ 実績（2021年）も 2/15〜4/22 に40夜＋前日引渡し2日で施工。"]:
    ws_g.cell(row=rr, column=2, value=note).font = f(8)
    rr += 1
ws_g.freeze_panes = "D8"
a4(ws_g, True, ws_g.PAPERSIZE_A3)
ws_g.page_setup.fitToHeight = 1

ws_w = wb.create_sheet("全体工程表")
ws_w["A1"] = "全体工程表（見積〜引渡し・週単位）"
ws_w["A1"].font = f(14, True)
ws_w.column_dimensions["A"].width = 4
ws_w.column_dimensions["B"].width = 50
W0 = dt.date(2026, 10, 19)
weeks = [W0 + dt.timedelta(weeks=i) for i in range(20)]
for j, w in enumerate(weeks):
    ws_w.column_dimensions[get_column_letter(3 + j)].width = 6
    c = ws_w.cell(row=3, column=3 + j, value=f"{w.month}/{w.day}")
    c.font = f(8); c.fill = HEAD; c.border = BORDER; c.alignment = Alignment(horizontal="center")
for i, h in enumerate(["No", "項目"], 1):
    c = ws_w.cell(row=3, column=i, value=h)
    c.font = f(9, True); c.fill = HEAD; c.border = BORDER
DD = dt.date
PHASES = [
    ("図面受領・積算・見積提出", DD(2026, 10, 21), DD(2026, 10, 30), BAR),
    ("見積調整・契約", DD(2026, 11, 2), DD(2026, 11, 13), BAR),
    ("実施設計図承認・施工承認図／製作図作成", DD(2026, 11, 9), DD(2026, 11, 27), BAR),
    ("分電盤 製作・納品（約5週）", DD(2026, 11, 16), DD(2026, 12, 25), BAR),
    ("建具・FR三方枠・スリット・サイン 工場製作", DD(2026, 11, 30), DD(2027, 1, 8), BAR),
    ("GU支給品（照明・床材・什器・サイン）納期調整", DD(2026, 11, 16), DD(2027, 1, 8), BAR3),
    ("施工承認申請（様式1〜9：着工10日前）", DD(2026, 12, 21), DD(2026, 12, 25), BAR2),
    ("着工届・作業員名簿・A/B工事引渡し確認（7日前）", DD(2027, 1, 5), DD(2027, 1, 8), BAR2),
    ("C工事 施工（40夜）", NIGHTS[0], NIGHTS[-1], BAR),
    ("内装監理室検査・消防検査", NIGHTS[36], NIGHTS[38], BAR2),
    ("引渡し・商品搬入・開店準備（GU様）", NIGHTS[-1], NIGHTS[-1] + dt.timedelta(days=6), BAR3),
    ("竣工図・竣工写真・シックハウス書類提出（1ヶ月以内）", NIGHTS[-1], NIGHTS[-1] + dt.timedelta(days=27), BAR2),
]
for k, (name, s, e, fill) in enumerate(PHASES, 1):
    rr = 3 + k
    for col, v in ((1, k), (2, name)):
        c = ws_w.cell(row=rr, column=col, value=v)
        c.font = f(9); c.border = BORDER
    for j, w in enumerate(weeks):
        c = ws_w.cell(row=rr, column=3 + j)
        c.border = BORDER
        if s <= w + dt.timedelta(days=6) and e >= w:
            c.fill = fill
ws_w.cell(row=5 + len(PHASES), column=2, value="※ 年末年始（12/28〜1/4）は館の工事休止を想定し、着工を1/12と仮定。").font = f(8)
a4(ws_w, True)
ws_w.page_setup.fitToHeight = 1
wb.save(PREFIX + "工程表_v2.xlsx")

# ===== 3. 拾い出し根拠 =====
wb = Workbook()
ws = wb.active
ws.title = "拾い出し根拠"
ws["A1"] = "拾い出し根拠資料（簡易版）"
ws["A1"].font = f(14, True)
ws["A2"] = f"根拠の区分 → 図面：2413区画 施工図一式から拾い／実績：2021/01 同物件見積の数量／図面・実績：両者でほぼ一致。単価は2021実績売価×{K}。"
ws["A2"].font = f(9)
head_row(ws, 4, ["工種", "名称", "数量", "単位", "根拠", "2021基準単価", "2026単価", "拾い出しメモ"], [18, 34, 9, 6, 9, 12, 11, 52])
r = 5
for sec, rows in ITEMS:
    for name, spec, q, u, base, src, memo in rows:
        for col, v in enumerate([sec[3:], name, q, u, src, base, price(base), memo], 1):
            c = ws.cell(row=r, column=col, value=v)
            c.font = f(9); c.border = BORDER
            c.alignment = Alignment(wrap_text=col in (2, 8), vertical="center")
        for col in (3, 6, 7):
            ws.cell(row=r, column=col).number_format = YEN
        r += 1
ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:H{r-1}"
a4(ws, True)

ws2 = wb.create_sheet("面積・台数の基本数量")
head_row(ws2, 1, ["項目", "数量", "単位", "出典"], [34, 10, 6, 60])
BASIC = [
    ("区画面積", 1064.0, "㎡", "照明器具表 床面積"),
    ("売場（レジ66.8・SW21.6含む）", 855.1, "㎡", "1064.0−BOH208.9"),
    ("BOH（FR含む）", 208.9, "㎡", "フロアマップ"),
    ("　FR／補正室／倉庫／備品庫／休憩室／店長室", 208.9, "㎡", "36.0／15.8／94.4／17.9／35.5／9.5"),
    ("天井高", 3750, "mm", "各図"),
    ("ファサード長さ", 20.8, "m", "ボーダー重量計算 6.790+7.283+6.745"),
    ("ダウンライト（BD・UD）", 396, "台", "照明器具表 BD1 99/BD3 87/BD4 12/BD5 59/BD6 5/UD6 37/UD7 79/UD8 18"),
    ("直付・スポット", 41, "台", "BL1 8/BL2 5/SP1 17/SP2 11"),
    ("コンセント", 87, "個", "ダブル41/抜止接地42/天井裏4"),
    ("POS/TEL/FAX アウトレット", 20, "個", "照明器具表"),
    ("分電盤", 2, "面", "盤図 電灯盤-1 MCB200AT ELB24回路／電灯盤-2 175AT ELB28回路"),
    ("FRブース", 10, "室", "小8＋大1＋HC1（実績見積 FR計10）"),
    ("建具", 7, "組", "WD-1A〜1D・WD-3A・WD-5A・SD-1A（＋FR WD-4A）"),
]
for i, row in enumerate(BASIC, 2):
    for col, v in enumerate(row, 1):
        c = ws2.cell(row=i, column=col, value=v)
        c.font = f(9); c.border = BORDER

ws3 = wb.create_sheet("v1からの変更点")
head_row(ws3, 1, ["区分", "内容"], [16, 100])
CHANGES = [
    ("支給品", "照明器具・床塩ビタイル材・カウンター類・LEDサインはGU支給。材料費を外し、取付・施工費のみ計上（v1比 約1,900万円減）"),
    ("造作", "レジバック・FR・ストック棚・流し台・店長室収納などの大型造作を削除。FRはGU標準の三方枠＋カーテンで計上"),
    ("読み違い修正", "化粧石膏ボードGB-DはBOH壁の仕上げ。BOH天井LGS・化粧ボードを削除し、天井は全面AEP（1,065㎡）"),
    ("追加", "コーナーガード67本、BOHクロス・化粧ボード、標準化項目（壁面スリット・誘導サイン・HCスツール等）、付帯（指定産廃・揚重・駐車場）、夜間警備"),
    ("共通費", "共通費率（%）方式をやめ、実績の諸経費内訳（現場管理費・監理・会社経費等）で計上"),
    ("単価", f"2021年実績売価×{K}（2021→2026 の労務・資材上昇を25%と仮定）"),
    ("要確認", "店頭サッシ撤去の工事区分、SW遮光シートの要否、ロールスクリーン（実績は追加工事）"),
]
for i, (a, b) in enumerate(CHANGES, 2):
    for col, v in enumerate((a, b), 1):
        c = ws3.cell(row=i, column=col, value=v)
        c.font = f(9); c.border = BORDER; c.alignment = Alignment(wrap_text=True, vertical="top")
wb.save(PREFIX + "拾い出し根拠_v2.xlsx")
print("saved")
