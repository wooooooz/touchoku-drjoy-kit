# -*- coding: utf-8 -*-
"""sheet_style.py  20260925 — 当直表xlsxの意匠（色・書体・罫線・寸法の正本） / 20260929
月曜はじまりの月間カレンダー（A4縦）用。幾何は duty_sheet.py、見た目はここ。

意匠の狙い（あとで直す人へ）
 1. 主役は氏名。氏名 14pt 黒（オンコールは同じ大きさの灰字）。それ以外の文字は全部 10pt 以下の灰にして、氏名より前に出ない。
 2. 紙を使い切る。列幅は A4縦の刷り幅（余白 0.35in・≒545pt）を 7 日で割って決め、行高は apply_page が
    週数×レーン数から逆算して埋める（レーン 18〜33pt）。レーンが多い科（9・13）は下限で止め、あとは fit で縮む。
 3. 罫線は薄灰 D9D9D9 の細線だけ。週の切れ目は日付行の上と備考行の下に中太線 595959 を1本。それ以外の太い線は無い。
 4. 左の枠ラベル・当オ・曜日見出しは塗らない（前版のチャコール帯は氏名より目立って主従が逆だった）。
 5. 土日祝の色は 日付セル（土=B9CEF2・日祝=F2B9B9）と、名前セル（その1/3の濃さ）。土と日祝は白との差を揃えてある。
    白黒コピーでは 6・7 列目という位置、太い数字の色の濃さ、祝日は祝日名の文字が手掛かりになる。
 6. 立たない日と前後月のマスは薄い斜線ハッチ（C0C0C0 の lightUp）。「斜線＝該当なし」の書類慣習。書いてよい空欄は真っ白。
 7. 日付は 13pt 太字で中央（氏名 14pt より一段軽い）。平日は 595959、土は 1F4E79、日祝は C00000。祝日名は 7.5pt で数字の下の行に出る。
 8. 太字は氏名では使わず「交代＝赤太字」のために空けてある。近接当直＝薄橙+橙細枠、未入力＝薄赤+赤点線。塗りと線の両方で白黒でも区別。
 9. タイトル 22pt 太字。1〜3行目は B〜J の結合セルに中央揃え（左揃えだと Excel の字詰め余白で表の左端と揃わない）。凡例は 9pt で折り返し 2 行。
10. フッタは付けない（掲示物に印刷日時とファイル名は要らない）。集計（印刷範囲外）は 9pt・薄罫線。

duty_sheet 側との約束
 - 曜日見出し = header_band（月〜金）/ header_frame（土日）、枠ラベル = header_frame、当 = header_slot、オn・備考 = header_band。
   いずれも塗り無し・罫線無し（border("header") は空）。
 - レーンのセルは fill("name") / fill("name_oncall") が白の塗りを返すので、土日の row_* は日付セルにだけ残る。
 - 日付行は border("day")（上=中太）、備考行は border("note")（下=中太）。前後月の空マスにも同じ role を当てると週の線が途切れない。
 - 祝日名は日付の TextBlock のあとに "\\n"+祝日名 で続ける（align("day") は wrap なので 2 行目に落ちる。半角空白だと語の途中で折れる）。
 - apply_page(ws, wide) は印刷設定に加えて、A列のキー（date / note / 枠キー）から週ブロックを数えて行高を決め直す（fit_rows）。
   列幅 COL_W は Normal スタイルが Calibri 11（openpyxl 既定）のままである前提。
"""
from __future__ import annotations

import math
from copy import copy as _copy

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.cell.text import InlineFont

__all__ = [
    "FONT_BODY", "FONT_TITLE", "BASE_SIZE", "SUFFIX_DELTA",
    "COL_W", "ROW_H", "font", "fill", "border", "align", "inline", "with_sides",
    "DXF", "PAGE", "LEGEND", "LEGEND_TOBAN", "C_TOBAN_BG", "C_TOBAN_INK", "apply_page", "fit_rows",
]

# ---------------------------------------------------------------- 書体・寸法
FONT_BODY = "BIZ UDPゴシック"        # 本文の書体名（掲示物向けUD書体。Win10 1809 以降の日本語環境に標準）
FONT_TITLE = "BIZ UDPゴシック"       # タイトルの書体名（同族。大きさと太さで差を付ける）
BASE_SIZE = 14.0                    # 氏名セルの基準サイズ(pt)
SUFFIX_DELTA = -2.0                 # 表示名の「足した部分」のサイズ差（固定・-2.0）

# Excel の文字幅（Normal=Calibri 11 で 1 文字幅 ≒ 7px、セルは W*7+5 px）
# A4縦の刷り幅 545pt ≒ 727px = 枠ラベル(110) + 当オ(33) + 7日×82px(11.1文字幅)
COL_W = {
    # --- カレンダー ---
    "cal_day_portrait": 11.1,   # 縦向きの曜日列。「丙川太郎」(14/12pt) が縮まずに入る 82px
    "cal_day": 17.0,            # 横向きの曜日列（横向きは高さで縮むので通常は使わない）
    "lane_label": 15.0,         # B列 枠ラベル「昼夜 セカンド以下」(10pt) が入る 110px
    "lane_kind": 4.0,           # C列 当／オn
    # --- 旧マトリクス互換・集計 ---
    "day": 4.5, "weekday": 12.0, "name": 10.0, "note": 14.0, "gap": 2.0,
    "summary_name": 9.0, "summary_num": 5.5,
}

# ポイント。Excel は 0.75pt（=1px）刻みで持つのでその倍数。cal_* は apply_page が週数×レーン数で決め直す
ROW_H = {
    "title": 34.5,              # 22pt 太字
    "subtitle": 15.0,           # 10pt
    "legend": 39.0,             # 9pt × 3 行（折り返し。3行目=職位の印の説明）
    "cal_wd": 18.0,             # 曜日見出し
    "cal_date": 30.0,           # 日付行。13pt の数字の下に 7.5pt の祝日名を折り返しても入る
    "cal_lane": 24.75,          # レーン行の名目値（4レーン×5週で紙を埋める値。apply_page が上書き）
    "cal_note": 12.0,           # 備考行。8pt 1 行ぶん。週の間の溝として働く
    # --- 旧マトリクス互換 ---
    "header_frame": 22.5, "header_band": 15.75, "header_slot": 17.25, "day": 22.5,
}

# 行高の逆算に使う値（apply_page / fit_rows）
LANE_H_MIN, LANE_H_MAX = 18.0, 45.0     # 14pt の氏名が欠けない下限／上限
                                        # 20260912: 33.0 だと枠が2本の科（歯科口腔外科）で紙の下 1/3 が空いた。
                                        # 壁に貼る紙なので、行を高くして埋めるほうが読みやすい
DATE_H_MIN, NOTE_H_MIN = 27.0, 12.0     # レーンが多くて下限に当たったときの日付行・備考行（祝日名2行目と 11pt の「備考」が欠けない下限）

# ---------------------------------------------------------------- 配色
C_INK = "000000"          # 氏名の黒
C_INK_SUB = "595959"      # オンコール・見出し語・凡例の灰（白黒コピーでも黒と区別がつく濃さ）
C_INK_DIM = "404040"      # 平日の日付・枠ラベル
C_MUTE = "A6A6A6"         # 立たないマスに何か書かれた場合の文字色
C_WHITE = "FFFFFF"

C_SAT = "B9CEF2"          # 土曜の日付セル（20260912: 日祝と同じ濃さに揃えた。
                          #  元の DCE6F2 は白との差が日祝の1/3しかなく、名前セルの
                          #  「1/3の濃さ」の色がほぼ白になって見えなかった）
C_SUN = "F2B9B9"          # 日祝の日付セル
C_SAT_INK = "1F4E79"      # 土の数字
C_SUN_INK = "C00000"      # 日祝の数字・祝日名

# 当番当直（管理者 20260925: 3か所で同じ見た目にする）: グレー地に白字。
#   色は DTX14「当直表 ファースト」の条件付き書式（当番当直シートを引く式）と同じ 595959。
#   アプリの画面（renderer の .slot select.tb）と dtx14_fill.check_toban（15時の反映の確かめ）も同じ値を使う
C_TOBAN_BG = "595959"
C_TOBAN_INK = C_WHITE
C_TOBAN_INK_OC = "D9D9D9" # 当番当直のマスのオンコール（開院記念日の昼など）。白の当直と見分けがつき、灰地でも読める

C_BLOCK_LINE = "C0C0C0"   # 斜線ハッチの線
C_BLOCK_BG = "F5F5F5"     # 斜線ハッチの地

C_RULE = "D9D9D9"         # マスの細線
C_WEEK = "595959"         # 週の切れ目の中太線
C_FRAME = "404040"        # 外枠（outer）

C_RED = "C00000"          # 交代・未入力
C_ORANGE = "ED7D31"       # 近接当直の枠線
C_ORANGE_BG = "F8CBAD"    # 近接当直の塗り
C_MISS_BG = "FCE4E4"      # 未入力の塗り（白紙の月は当直1のマスが全部これになるので薄く）
C_HEAD_3 = "E6E6E6"       # 集計の見出し

# ---------------------------------------------------------------- role 表
# (書体, サイズ, 太字, 文字色)
_FONTS: dict[str, tuple[str, float, bool, str]] = {
    "title":         (FONT_TITLE, 22.0, True,  C_INK),
    "subtitle":      (FONT_BODY,  10.0, False, C_INK_SUB),
    "legend":        (FONT_BODY,   9.0, False, C_INK_SUB),
    "header_frame":  (FONT_BODY,  14.0, True,  C_INK_DIM),   # 土日の曜日見出し・枠ラベル
    "header_band":   (FONT_BODY,  11.0, False, C_INK_SUB),   # 月〜金の曜日見出し・オn・備考
    "header_slot":   (FONT_BODY,  11.0, True,  C_INK),       # 当
    "day":           (FONT_BODY,  13.0, True,  C_INK_SUB),   # 平日の日付（氏名より一段軽く）
    "weekday":       (FONT_BODY,  11.0, False, C_INK),
    "weekday_sat":   (FONT_BODY,  13.0, True,  C_SAT_INK),   # 土の日付
    "weekday_sun":   (FONT_BODY,  13.0, True,  C_SUN_INK),   # 日祝の日付
    "holiday_name":  (FONT_BODY,   7.5, False, C_SUN_INK),   # 6文字の祝日名が 82px の日付セルに 1 行で入る
    "name":          (FONT_BODY, BASE_SIZE,                False, C_INK),
    "name_suffix":   (FONT_BODY, BASE_SIZE + SUFFIX_DELTA, False, C_INK),
    "name_oncall":   (FONT_BODY, BASE_SIZE,                False, C_INK_SUB),
    "name_oncall_suffix": (FONT_BODY, BASE_SIZE + SUFFIX_DELTA, False, C_INK_SUB),
    "name_toban":    (FONT_BODY, BASE_SIZE,                False, C_TOBAN_INK),     # 当番当直（当直は duty_sheet が太字にする）
    "name_toban_oncall": (FONT_BODY, BASE_SIZE,            False, C_TOBAN_INK_OC),
    "note":          (FONT_BODY,   8.0, False, C_INK_SUB),
    "blocked":       (FONT_BODY,   8.0, False, C_MUTE),
    "summary_head":  (FONT_BODY,   9.0, True,  C_INK),
    "summary_body":  (FONT_BODY,   9.0, False, C_INK),
}

_WHITE_FILL = PatternFill("solid", fgColor=C_WHITE, bgColor=C_WHITE)

# 塗らない role は None
_FILLS: dict[str, PatternFill | None] = {
    "header_frame": None,                       # 見出し・枠ラベルは塗らない
    "header_band":  None,
    "header_slot":  None,
    "row_weekday":  None,
    "row_sat":      PatternFill("solid", fgColor=C_SAT, bgColor=C_SAT),
    "row_sun":      PatternFill("solid", fgColor=C_SUN, bgColor=C_SUN),
    # レーンのセルは白で塗る = duty_sheet の `fill(role) or 行の塗り` で土日の色がレーンに広がらない
    "name":         _WHITE_FILL,
    "name_input":   _WHITE_FILL,
    "name_oncall":  _WHITE_FILL,
    "name_toban":   PatternFill("solid", fgColor=C_TOBAN_BG, bgColor=C_TOBAN_BG),   # 当番当直のマス（20260925）
    "name_toban_oncall": PatternFill("solid", fgColor=C_TOBAN_BG, bgColor=C_TOBAN_BG),
    # 斜線ハッチ: patternType≠solid では fgColor=線の色・bgColor=地の色
    "blocked":      PatternFill(patternType="lightUp", fgColor=C_BLOCK_LINE, bgColor=C_BLOCK_BG),
    "note":         None,
    "summary_head": PatternFill("solid", fgColor=C_HEAD_3, bgColor=C_HEAD_3),
    "summary_body": None,
}

_S_R = Side(style="thin", color=C_RULE)
_S_WK = Side(style="medium", color=C_WEEK)
_S_M = Side(style="medium", color=C_FRAME)

_BORDERS: dict[str, Border] = {
    "header":     Border(),                                                  # 見出し・ラベルは線なし
    "name":       Border(left=_S_R, right=_S_R, top=_S_R, bottom=_S_R),
    "day":        Border(left=_S_R, right=_S_R, top=_S_WK, bottom=_S_R),      # 上=週の切れ目
    "note":       Border(left=_S_R, right=_S_R, top=_S_R, bottom=_S_WK),      # 下=週の切れ目
    "summary":    Border(left=_S_R, right=_S_R, top=_S_R, bottom=_S_R),
    "outer":      Border(left=_S_M, right=_S_M, top=_S_M, bottom=_S_M),
    "none":       Border(),
    "frame_left": Border(left=_S_M, right=_S_R, top=_S_R, bottom=_S_R),
}

_ALIGNS: dict[str, Alignment] = {
    "title":    Alignment(horizontal="center", vertical="center"),
    "subtitle": Alignment(horizontal="center", vertical="center"),
    "legend":   Alignment(horizontal="center", vertical="center", wrap_text=True),
    "header":   Alignment(horizontal="center", vertical="center", shrink_to_fit=True),
    "day":      Alignment(horizontal="center", vertical="center", wrap_text=True),
    "weekday":  Alignment(horizontal="center", vertical="center", shrink_to_fit=True),
    "name":     Alignment(horizontal="center", vertical="center", shrink_to_fit=True),
    "note":     Alignment(horizontal="left",   vertical="center", shrink_to_fit=True),
    "summary":  Alignment(horizontal="center", vertical="center"),
}


# ---------------------------------------------------------------- API
def font(role: str) -> Font:
    """role に応じた Font。未知の role は本文書体・基準サイズの素の Font。"""
    name, size, bold, color = _FONTS.get(role, (FONT_BODY, BASE_SIZE, False, C_INK))
    return Font(name=name, size=size, bold=bold, color=color)


def inline(role: str) -> InlineFont:
    """リッチテキスト（CellRichText / TextBlock）用。font(role) と同じ見た目の InlineFont。"""
    f = font(role)
    return InlineFont(rFont=f.name, sz=f.size, b=f.bold, i=f.italic, u=f.underline, color=f.color)


def fill(role: str) -> PatternFill | None:
    """role に応じた PatternFill。塗らない role と未知の role は None。"""
    src = _FILLS.get(role)
    return _copy(src) if src is not None else None


def border(role: str) -> Border:
    """role に応じた Border。未知の role は罫線なし。"""
    src = _BORDERS.get(role)
    return _copy(src) if src is not None else Border()


def align(role: str) -> Alignment:
    """role に応じた Alignment。未知の role は中央揃え。"""
    src = _ALIGNS.get(role)
    return _copy(src) if src is not None else Alignment(horizontal="center", vertical="center")


def with_sides(base: Border, **sides: Side | None) -> Border:
    """base の一部の辺だけ差し替える。
    例: with_sides(border("name"), left=border("outer").left)
    """
    kw = dict(left=base.left, right=base.right, top=base.top, bottom=base.bottom,
              diagonal=base.diagonal, diagonal_direction=base.diagonal_direction,
              outline=base.outline, vertical=base.vertical, horizontal=base.horizontal)
    kw.update(sides)
    return Border(**kw)


# ---------------------------------------------------------------- 条件付き書式
def _box(side: Side) -> Border:
    return Border(left=side, right=side, top=side, bottom=side)


# dxf の塗りは bgColor が要る（fgColor だけだと Excel が無視する）ので両方に同じ色を入れる
DXF = {
    # 交代（当初と違う人）: 赤字・太字。氏名は通常太字にしないので白黒コピーでも判る
    "changed": DifferentialStyle(font=Font(bold=True, color=C_RED)),
    # 連続/近接当直の警告: 薄橙の塗り + 橙の細枠（レーンは白なので塗りだけでも白黒で出る）
    "crowded": DifferentialStyle(
        fill=PatternFill("solid", fgColor=C_ORANGE_BG, bgColor=C_ORANGE_BG),
        border=_box(Side(style="thin", color=C_ORANGE)),
    ),
    # 必須なのに未入力: 薄赤 + 赤の点線枠（「ここに書く」の輪郭）
    "missing": DifferentialStyle(
        fill=PatternFill("solid", fgColor=C_MISS_BG, bgColor=C_MISS_BG),
        border=_box(Side(style="dashed", color=C_RED)),
    ),
    # 当番当直のマス（20260925）。そのマスだけに、ほかの規則より先に掛ける:
    #   オンコールの灰字（595959）は灰地に沈むので薄い灰にする／未入力でも灰地のまま赤の点線枠だけ出す／
    #   交代は赤地に白字（管理者 20260929。以前は白字＋赤の枠。DTX14・アプリの画面と同じ）
    "toban_oncall": DifferentialStyle(font=Font(color=C_TOBAN_INK_OC)),
    "toban_changed": DifferentialStyle(font=Font(bold=True, color=C_TOBAN_INK),
                                       fill=PatternFill("solid", fgColor=C_RED, bgColor=C_RED)),
    "toban_missing": DifferentialStyle(
        fill=PatternFill("solid", fgColor=C_TOBAN_BG, bgColor=C_TOBAN_BG),
        border=_box(Side(style="dashed", color=C_RED)),
    ),
}

# ---------------------------------------------------------------- 印刷体裁
FOOTER = {"left": "印刷 &D &T", "right": "&F", "size": 8}   # header_footer を有効にしたときの中身

PAGE = {
    "cal_orientation": "portrait",     # カレンダーの既定は縦
    "orientation_default": "portrait",
    "orientation_wide": "landscape",
    "wide_threshold": 10,
    "fit_to_width": 1,
    "fit_to_height": 1,
    "paper_size": 9,                   # A4
    # インチ。掲示物なので余白は最小限（プリンタの機械的な余白 ≒0.2in を下回らない）
    "margins": {"left": 0.35, "right": 0.35, "top": 0.4, "bottom": 0.4, "header": 0.2, "footer": 0.2},
    "header_footer": None,             # 掲示物にフッタは付けない（付けるなら "&L印刷 &D &T&R&F"）
    "print_title_rows": None,          # カレンダーは1枚
    "show_gridlines": False,
    "zoom": 100,
    # 余白を除いた刷れる高さ(pt)。fit_rows の逆算に使う。Excel 実測で補正する
    # 名目 784/527pt だが Excel は行高の名目値の約 92.6% で刷る（PNG 実測）ので名目値に換算してある
    "print_height_pt": {"portrait": 848.0, "landscape": 570.0},
}

# 行3 の凡例。B〜J の結合セルで 9pt・折り返し 2 行に収まる長さ。斜線・空欄・赤字を先頭に
LEGEND = ("赤字＝当初から交代した人　灰字＝オンコール（旧表の括弧書き）　"
          "教＝教員　研＝大学院生　専＝専攻医　初＝初期研")
LEGEND_TOBAN = "　グレー地に白字＝当番当直"   # 当番当直のマスがある表だけ凡例の末尾に足す（20260925）


def _calendar_rows(ws):
    """A列のキーから週ブロックの行を拾う → (先頭キー行, 日付行[], レーン行[], 備考行[])。"""
    first, dates, lanes, notes = None, [], [], []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        k = str(v).strip() if v is not None else ""
        if not k:
            continue
        if first is None:
            first = r
        if k == "date":
            dates.append(r)
        elif k == "note":
            notes.append(r)
        elif k.count("|") >= 3:
            lanes.append(r)
    return first, dates, lanes, notes


def _snap(pt: float) -> float:
    return math.floor(pt / 0.75) * 0.75


def fit_rows(ws, wide: bool = False):
    """週数×レーン数から行高を決め直し、A4 1 枚を使い切る。
    レーンが多くて下限に当たる科は、日付行・備考行も詰めたうえで fit-to-page の縮小に任せる。
    戻り値は {lane_h, date_h, note_h, total_pt, scale}。カレンダーでないシートは None。"""
    first, dates, lanes, notes = _calendar_rows(ws)
    if not dates or not lanes:
        return None
    page_h = PAGE["print_height_pt"]["landscape" if wide else "portrait"]
    fixed = 0.0
    block = set(dates) | set(lanes) | set(notes)
    last = max(block)
    # 週ブロックの前（表題）と後ろ（凡例）に置いた行も、紙を使い切る計算に入れる
    for r in list(range(1, first)) + [r for r in range(first, last + 2) if r not in block]:
        rd = ws.row_dimensions[r]
        if not rd.hidden and r not in block:
            fixed += rd.height if rd.height is not None else 15.0
    nb, nl = len(dates), len(lanes) / len(dates)
    date_h, note_h = ROW_H["cal_date"], ROW_H["cal_note"]
    lane_h = (page_h - fixed - nb * (date_h + note_h)) / (nb * nl)
    if lane_h < LANE_H_MIN:
        date_h, note_h = DATE_H_MIN, NOTE_H_MIN
        lane_h = (page_h - fixed - nb * (date_h + note_h)) / (nb * nl)
    lane_h = max(LANE_H_MIN, min(LANE_H_MAX, _snap(lane_h)))
    for r in dates:
        ws.row_dimensions[r].height = date_h
    for r in lanes:
        ws.row_dimensions[r].height = lane_h
    for r in notes:
        ws.row_dimensions[r].height = note_h
    total = fixed + nb * (date_h + note_h) + len(lanes) * lane_h
    return {"lane_h": lane_h, "date_h": date_h, "note_h": note_h,
            "total_pt": total, "scale": min(1.0, page_h / total)}


def apply_page(ws, wide: bool = False) -> None:
    """印刷・表示の設定を一括で入れ、カレンダーなら行高を紙に合わせて決め直す。
    印刷範囲は幾何なので呼ぶ側で。"""
    from openpyxl.worksheet.properties import PageSetupProperties

    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = PAGE["orientation_wide"] if wide else PAGE["orientation_default"]
    ws.page_setup.paperSize = PAGE["paper_size"]
    ws.page_setup.fitToWidth = PAGE["fit_to_width"]
    ws.page_setup.fitToHeight = PAGE["fit_to_height"]
    m = PAGE["margins"]
    ws.page_margins.left = m["left"]
    ws.page_margins.right = m["right"]
    ws.page_margins.top = m["top"]
    ws.page_margins.bottom = m["bottom"]
    ws.page_margins.header = m["header"]
    ws.page_margins.footer = m["footer"]
    ws.print_options.horizontalCentered = True
    if PAGE["header_footer"] is not None:
        ws.oddFooter.left.text = FOOTER["left"]
        ws.oddFooter.left.size = FOOTER["size"]
        ws.oddFooter.right.text = FOOTER["right"]
        ws.oddFooter.right.size = FOOTER["size"]
    if PAGE["print_title_rows"]:
        ws.print_title_rows = PAGE["print_title_rows"]
    ws.sheet_view.showGridLines = PAGE["show_gridlines"]
    ws.sheet_view.zoomScale = PAGE["zoom"]
    fit_rows(ws, wide)
