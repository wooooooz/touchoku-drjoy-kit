# -*- coding: utf-8 -*-
"""admin_gui.py — 当直勤務表 管理ツール（事務担当用・GUI・exe）。20261013

4タブ（20260925 管理者の決定で②④を作り替えた）:
  ① 月フォルダを作る       … 年月（既定=翌月、25日以降は翌々月）と締切（前の月の20日 12:00＝new_month.deadline_for を
                              月を変えるたびに入れ直す）→ month_setup.create。チェックが入っていれば SharePoint の DTX14 の雛形も
                              無いときだけ作る（new_month.make_dtx14）。ふだんは毎月1〜5日に15時の反映が自動で作るので、作り直し用。
  ② 提出状況を確かめる     … マスターを読んで、未提出の科・「当直のみ」・名簿外の医師・重複/弾いた行 をログ欄に出し、
                              集約\\集約レポート.txt に保存（aggregate report_only。互換3シート・duty JSON は作らない）
  ③ この月を確定           … マスターの confirmed を立て、各科の baseline に そのときの duties を焼く
                              （master_store.confirm。20260912 に config\\確定.txt から移した）／解除
  ④ 手入力の部署を取り込む … 看護部・薬剤部などから届いた DTX14 の形の Excel の、その部署の行だけを正本 DTX14 に写す
                              （dtx14_manual_import）。以前の「DTX14 に書き出す」は 15時の反映が毎日するので外した。
本番の場所は contract.PROD_ROOT（共有）。exe（--windowed）では print を画面のログ欄へ流す。

20260925（公開前の見直し）:
  - ②③の月の既定は「締切が過ぎて、まだ確定していない月」（無ければ来月以降でいちばん近い月）。以前は一番先の月
    （例 R8年12月）が選ばれていて、違う月を確定しかねなかった。月は「R8年11月分（2026年11月）」の形で出す。
  - ③の確定は、月を大きな字で出し、まだ提出が無い科の数と名前を出す。まだの科があれば、もう一度確かめる。
    解除にも確かめの窓を出す（解除で元に戻せるのは、まだ交代の承認が1件も無いときだけ）。
  - ④はファイルごとに取り込む部署を選ぶ（既定＝そのファイルで空でなく、正本と違う部署）。14:55〜15:15 は取り込まない（20261010 に縮めた）。
  - 隠しの確かめ: admin_tool.exe --selfcheck 結果.txt [--com] … 使うモジュールを全部読み込み、同梱の data を読んで終わる
    （exe のビルドの確かめ用。--com は Excel で日付を1つ読む＝win32timezone が同梱されているか）。

20260926:
  - 「確定しているか」は確定の正本（集約\\確定.json）で決める（month_confirmed＝master_store.effective_confirmed）。
    ②③の月の「確定済み」の印・③の二重押しの確かめ・④の月の既定も同じ。確定.json が壊れていれば③の確定は止める（解除は通す＝作り直す）。
  - ①の説明: すでにある月でもアプリ（起動係・zip）は新しい版に置き換わる（全科の予定は触らない）と正しく書いた。
    正本の DTX14 が2本あるときは「まだ無い」ではなく「2本あります。1本にしてください」と出す。
  - ④の月の既定も②③と同じ default_month。
"""
import contextlib
import datetime as dt
import io
import os
import queue
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402
import master_store as MST    # noqa: E402
import month_setup as MS      # noqa: E402
import new_month as NM        # noqa: E402
import dtx14_manual_import as MI   # noqa: E402


def use_my_dtx_root():
    """SharePoint の「当直勤務表_正本」を、この PC で同期している場所に向ける（15時の反映は管理者の PC の場所を決め打ちで持つ）。
    戻り: 使う場所（見つからなければ None）。"""
    import dtx14_daily_sync as S
    root = MI.dtx_root()
    if root:
        S.DTX_ROOT = root
    return root


FONT = "Yu Gothic UI"        # 20260925 管理者: BIZ UDPゴシックは字が詰まって見にくい → Yu Gothic UI。老眼でも読める大きさに
FONT_SIZE = 12


# ---------------------------------------------------------------- 月の既定（②③。20260925）
def _sync_from():
    try:
        import dtx14_daily_sync as S
        return tuple(S.SYNC_FROM)
    except Exception:
        return (2026, 11)


def parse_deadline(s, y, m):
    """config の締切の文字（'10月20日(火) 12:00'）→ datetime。年は対象月 (y, m) から補う（締切は対象月より前）。
    時刻が無ければその日の終わり。読めなければ None。"""
    mt = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日", s or "")
    if not mt:
        return None
    mm, dd = int(mt.group(1)), int(mt.group(2))
    hm = re.search(r"(\d{1,2})\s*[:：]\s*(\d{2})", s[mt.end():])
    yy = y if mm <= m else y - 1
    try:
        return dt.datetime(yy, mm, dd, int(hm.group(1)) if hm else 23, int(hm.group(2)) if hm else 59)
    except ValueError:
        return None


def month_deadline(y, m):
    """その月の締切（いまの config の締切。無い・読めなければ 前の月の20日 12:00＝new_month.deadline_for）。"""
    try:
        import json
        with open(os.path.join(MS.month_dir(y, m), "config", "current.json"), encoding="utf-8") as f:
            d = parse_deadline(str(json.load(f).get("deadline") or ""), y, m)
        if d:
            return d
    except (OSError, ValueError):
        pass
    try:
        return parse_deadline(NM.deadline_for(y, m), y, m)
    except Exception:
        return None


def _month_end(y, m):
    return dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)


def default_month(months, now=None, deadline_of=month_deadline, sync_from=None):
    """②③で最初に選んでおく月。months = [(年, 月, 確定済みか)]。
    1) アプリで回す月（15時の反映が書く月＝SYNC_FROM 以降）のうち、締切が過ぎて・まだ確定しておらず・その月が終わっていない月
       （いちばん早いもの）
    2) 無ければ、来月以降でいちばん近い月　3) それも無ければ一番最後の月。戻り: (年, 月) か None"""
    now = now or dt.datetime.now()
    sync_from = sync_from or _sync_from()
    live = [(y, m, c) for y, m, c in months if (y, m) >= sync_from] or list(months)
    if not live:
        return None
    due = []
    for y, m, c in live:
        if c or _month_end(y, m) < now.date():
            continue
        d = deadline_of(y, m)
        if d and d <= now:
            due.append((y, m))
    if due:
        return min(due)
    nxt = (now.year + (now.month == 12), now.month % 12 + 1)
    later = [(y, m) for y, m, _c in live if (y, m) >= nxt]
    if later:
        return min(later)
    return live[-1][:2]


def test_dirs():
    """試験用の月フォルダ {(年, 月): パス}（R8年N月(動作確認用テストフォルダ)。同じ月のふつうのフォルダが無いものだけ）。
    20261001 管理者: 9月・10月は試験用の名前にしたので、②③で選べるようにする（確定して交代申請を試す）。"""
    import set_acl as SA
    out = {}
    try:
        for name in os.listdir(C.PROD_ROOT):
            ym = SA.parse_test_dir(name)
            if ym and os.path.isdir(os.path.join(C.PROD_ROOT, name, "config")) and not os.path.isdir(MS.month_dir(*ym)):
                out[ym] = os.path.join(C.PROD_ROOT, name)
    except OSError:
        pass
    return out


def month_dir(y, m):
    """その月のフォルダ（ふつうのフォルダが無く試験用があれば試験用）。"""
    d = MS.month_dir(y, m)
    if os.path.isdir(d):
        return d
    return test_dirs().get((y, m), d)


def master_json(y, m):
    """ふつうの月は month_setup.master_json のまま。ふつうのフォルダが無く試験用があるときだけ試験用のマスター。"""
    if not os.path.isdir(MS.month_dir(y, m)):
        td = test_dirs().get((y, m))
        if td:
            return MST.master_path(td, y, m)
    return MS.master_json(y, m)


def month_confirmed(y, m):
    """その月の確定（確定の正本 集約\\確定.json で決める＝master_store.effective_confirmed。20260926）。
    戻り: (confirmed の dict か None, マスター, 問題の一覧)。マスターが無ければ (None, None, [])。
    確定.json が壊れていれば MST.TruthError、マスターが読めなければ例外。"""
    p = master_json(y, m)
    if not os.path.isfile(p):
        return None, None, []
    mst = MST.load(p, y, m)
    conf, probs = MST.effective_confirmed(p, mst)
    return conf, mst, probs


def is_confirmed(y, m):
    """一覧の「確定済み」の印（読めなければ False）。"""
    try:
        return bool(month_confirmed(y, m)[0])
    except Exception:
        return False


def month_label(y, m, closed=False, test=False):
    return (f"R{y - 2018}年{m}月分（{y}年{m}月）" + ("（動作確認用テストフォルダ）" if test else "")
            + ("　確定済み" if closed else ""))


def official_hits(y, m):
    """その月のフォルダにある、正本の名前の規則に合う DTX14（名前の一覧。読めなければ []）。"""
    try:
        import dtx14_daily_sync as S
        d = MI.month_folder(y, m)
        return sorted(n for n in os.listdir(d)
                      if n.lower().endswith(".xlsm") and not n.startswith("~$") and S.is_master_name(n, y, m))
    except Exception:
        return []


def missing_owners(mst, frames=None):
    """まだ提出が無い提出単位（マスターに無い・予定が空）。並びは枠マスタの順（aggregate の「まだ提出が無い科」と同じ単位）。"""
    frames = C.load_frames() if frames is None else frames
    depts = (mst or {}).get("depts") or {}
    return [o for o in C.owners_of(frames) if not (depts.get(o) or {}).get("duties")]


def _dpi_aware():
    """Windows の表示の拡大（125%・150%）で文字がぼやけないように、拡大に対応していると宣言する。
    宣言しないと Windows が画面を画像として引き伸ばすので、にじむ（20260925 管理者「ボケてる」）。Tk を作る前に呼ぶ。"""
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)      # モニターごとの拡大に対応
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self):
        _dpi_aware()
        super().__init__()
        self.title("当直勤務表 管理ツール（月フォルダ・提出状況・確定・手入力の部署）")
        # 拡大率に合わせて窓の大きさを決める（画面からはみ出さないように）
        f = max(1.0, self.winfo_fpixels("1i") / 96.0)
        w = min(int(1000 * f), int(self.winfo_screenwidth() * 0.92))
        h = min(int(880 * f), int(self.winfo_screenheight() * 0.88))
        self.geometry(f"{w}x{h}")
        self.minsize(min(int(760 * f), w), min(int(600 * f), h))
        import tkinter.font as tkfont
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont", "TkTooltipFont"):
            try:
                tkfont.nametofont(name).configure(family=FONT, size=FONT_SIZE)
            except Exception:
                pass
        style = ttk.Style(self)
        for k in ("TLabel", "TButton", "TEntry", "TCheckbutton", "TCombobox", "TLabelframe.Label"):
            style.configure(k, font=(FONT, FONT_SIZE))
        style.configure("TNotebook.Tab", font=(FONT, FONT_SIZE, "bold"), padding=(16, 8))
        self.option_add("*TCombobox*Listbox.font", (FONT, FONT_SIZE))
        ttk.Label(self, text=f"本番: {C.PROD_ROOT}", foreground="#666", font=(FONT, 11)).pack(anchor="w", padx=12, pady=(8, 0))
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(6, 0))
        self.t1 = ttk.Frame(nb, padding=12)
        self.t2 = ttk.Frame(nb, padding=12)
        self.t3 = ttk.Frame(nb, padding=12)
        self.t4 = ttk.Frame(nb, padding=12)
        nb.add(self.t1, text=" ① 月フォルダを作る ")
        nb.add(self.t2, text=" ② 提出状況を確かめる ")
        nb.add(self.t3, text=" ③ この月を確定 ")
        nb.add(self.t4, text=" ④ 手入力の部署を取り込む ")
        self._build1()
        self._build2()
        self._build3()
        self._build4()
        logf = ttk.LabelFrame(self, text="結果・ログ", padding=6)
        logf.pack(fill="both", expand=True, padx=10, pady=10)
        self.log = tk.Text(logf, height=12, font=(FONT, FONT_SIZE), wrap="word", state="disabled")
        sb = ttk.Scrollbar(logf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        nb.bind("<<NotebookTabChanged>>", lambda e: (self.refresh_months(), self.refresh4()))

    # ---- 共通
    def logw(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def run_bg(self, fn, btn, after=None):
        btn.configure(state="disabled")
        q = queue.Queue()

        def job():
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
                    opened = fn()
                q.put(("log", buf.getvalue()))
                if opened:
                    q.put(("open", opened))
            except SystemExit as e:
                q.put(("log", buf.getvalue() + f"\n中止: {e}\n"))
                q.put(("err", ("中止", str(e))))
            except Exception as e:
                import traceback
                q.put(("log", buf.getvalue() + f"\nエラー: {e}\n{traceback.format_exc()}\n"))
                q.put(("err", ("エラー", str(e))))
            finally:
                q.put(("done", None))

        def poll():
            done = False
            try:
                while True:
                    kind, payload = q.get_nowait()
                    if kind == "log":
                        self.logw(payload)
                    elif kind == "open":
                        try:
                            os.startfile(payload)
                        except Exception:
                            pass
                    elif kind == "err":
                        messagebox.showerror(*payload)
                    elif kind == "done":
                        done = True
            except queue.Empty:
                pass
            if done:
                btn.configure(state="normal")
                if after:
                    after()
            else:
                self.after(100, poll)

        threading.Thread(target=job, daemon=True).start()
        self.after(100, poll)

    def months(self):
        """[(年, 月, 確定済みか)]。確定済みかは確定の正本で決め直す（month_setup.is_closed はマスターの confirmed を見るため。20260926）。"""
        try:
            ms = [(y, m, is_confirmed(y, m)) for y, m, _c in MS.list_months()]
            ms += [(y, m, is_confirmed(y, m)) for (y, m) in sorted(test_dirs())]     # 試験用（20261001）
            return sorted(ms)
        except Exception:
            return []

    def refresh_months(self):
        """②③の月の一覧を作り直す。選んでいた月はそのまま（確定の印が変わっても）。
        まだ選んでいなければ default_month（締切が過ぎて、まだ確定していない月。20260925）。"""
        ms = self.months()
        tests = test_dirs()
        self._months23 = {month_label(y, m, c, (y, m) in tests): (y, m) for y, m, c in ms}
        vals = list(self._months23)
        real = [x for x in ms if x[:2] not in tests] or ms      # 試験用の月は既定にしない（20261001）
        try:
            dflt = default_month(real)
        except Exception:
            dflt = real[-1][:2] if real else None
        for cb in [c for c in (getattr(self, "cb2", None), getattr(self, "cb3", None)) if c is not None]:
            cur = getattr(cb, "_ym_sel", None)
            cb["values"] = vals
            pick = cur if cur in self._months23.values() else dflt
            lab = next((k for k, v in self._months23.items() if v == pick), "")
            cb.set(lab)
            cb._ym_sel = pick if lab else None

    def _cb_ym(self, cb):
        """コンボボックスで選んでいる (年, 月)（無ければ None）。"""
        ym = getattr(self, "_months23", {}).get(cb.get())
        cb._ym_sel = ym
        return ym

    def _bind_month(self, cb):
        cb.bind("<<ComboboxSelected>>", lambda e: self._cb_ym(cb))

    # ---- ① 月フォルダ
    def _build1(self):
        f = self.t1
        y, m = MS.default_target()
        self.y1, self.m1 = tk.StringVar(value=str(y)), tk.StringVar(value=str(m))
        self.dl1 = tk.StringVar()
        self.dtx1 = tk.BooleanVar(value=True)
        ttk.Label(f, text="月フォルダは、毎月1〜5日に自動で作られます（15時の自動処理が、次の次の月の分として\n"
                          "アプリの月フォルダと SharePoint の DTX14 の雛形を作ります。例: 10月1日に12月分）。\n"
                          "このタブは、作り直したいときや、自動で作られていなかったときだけ使ってください。",
                  foreground="#b35c00", justify="left").pack(anchor="w", pady=(0, 8))
        row = ttk.Frame(f)
        row.pack(anchor="w", pady=4)
        ttk.Label(row, text="対象").pack(side="left")
        ttk.Entry(row, textvariable=self.y1, width=6).pack(side="left", padx=(6, 2))
        ttk.Label(row, text="年").pack(side="left")
        ttk.Entry(row, textvariable=self.m1, width=4).pack(side="left", padx=(8, 2))
        ttk.Label(row, text="月分").pack(side="left")
        ttk.Label(row, text="（既定＝翌月。25日以降は翌々月）", foreground="#666").pack(side="left", padx=8)
        row2 = ttk.Frame(f)
        row2.pack(anchor="w", pady=4, fill="x")
        ttk.Label(row2, text="締切の表示").pack(side="left")
        ttk.Entry(row2, textvariable=self.dl1, width=30).pack(side="left", padx=6)
        ttk.Label(row2, text="既定は前の月の20日 12:00（土日祝なら前の平日）",
                  foreground="#666").pack(side="left")
        ttk.Checkbutton(f, text="SharePoint の DTX14 の雛形も作る（無いときだけ。あるものは上書きしません）",
                        variable=self.dtx1).pack(anchor="w", pady=(6, 0))
        self.st1 = ttk.Label(f, text="", foreground="#666", justify="left")
        self.st1.pack(anchor="w", pady=(2, 0))
        ttk.Label(f, text="押すと共有に「R8年N月」フォルダを作ります（アプリ・当直表ツール・名簿つき config・空のマスター・集約\\）。\n"
                          "すでにあれば config・当直表ツール・お読みくださいを作り直し、アプリ（起動係と本体の zip）は新しい版があれば置き換えます。\n"
                          "全科の予定（マスター）は触りません。\n"
                          "DTX14 の雛形は、前の月の正本を写して日付を送り、当直のマスを空にして作ります（Excel を使うので1分ほどかかります）。",
                  foreground="#444", justify="left").pack(anchor="w", pady=(8, 4))
        self.b1 = ttk.Button(f, text="月フォルダを作る", command=self.do1)
        self.b1.pack(anchor="e", pady=(8, 0))
        self.y1.trace_add("write", lambda *_: self._month1_changed())
        self.m1.trace_add("write", lambda *_: self._month1_changed())
        self._month1_changed()

    def _ym1(self):
        try:
            y, m = int(self.y1.get()), int(self.m1.get())
        except ValueError:
            return None
        return (y, m) if 2020 <= y <= 2100 and 1 <= m <= 12 else None

    def _month1_changed(self):
        """月を変えたら締切を入れ直し、DTX14 の雛形があるかを出す。"""
        ym = self._ym1()
        if not ym:
            return
        y, m = ym
        self.dl1.set(self._deadline1(y, m))
        try:
            p = MI.official_path(y, m)
            st = f"SharePoint の DTX14: あり（{os.path.basename(p)}）"
        except MI.ManualImportError:
            hits = official_hits(y, m)
            if len(hits) >= 2:
                st = f"SharePoint の DTX14: 正本が{len(hits)}本あります（{'、'.join(hits)}）。1本にしてください"
            else:
                st = (f"SharePoint の DTX14: まだ無い（R{y - 2018}.{m}月 当直勤務表）" if MI.dtx_root()
                      else "SharePoint の同期フォルダ（当直勤務表_正本）がこの PC に見つかりません")
        self.st1.configure(text=st)

    @staticmethod
    def _deadline1(y, m):
        """締切の初期値: すでにある月はいまの config の締切（作り直しで変えてしまわないように）、無ければ
        前の月の20日 12:00（土日祝なら前の平日。new_month.deadline_for＝自動で作るときと同じ）。"""
        try:
            import json
            with open(os.path.join(MS.month_dir(y, m), "config", "current.json"), encoding="utf-8") as f:
                cur = str(json.load(f).get("deadline") or "")
            if cur:
                return cur
        except (OSError, ValueError):
            pass
        try:
            return NM.deadline_for(y, m)
        except Exception:
            return ""

    def do1(self):
        ym = self._ym1()
        if not ym:
            messagebox.showerror("入力", "年・月を数字で入れてください")
            return
        y, m = ym
        dl = self.dl1.get().strip()
        with_dtx = self.dtx1.get()
        self.clear_log()

        def job():
            errs = []
            try:
                MS.create(y, m, dl or None)
            except (Exception, SystemExit) as e:        # 片方の失敗でもう片方を止めない（new_month.ensure と同じ）
                print(f"✗ アプリの月フォルダ: {e}")
                errs.append(f"アプリの月フォルダ: {e}")
            if with_dtx:
                print("")
                if not use_my_dtx_root():
                    errs.append("SharePoint の同期フォルダ（当直勤務表_正本）がこの PC に見つかりません")
                else:
                    try:
                        NM.make_dtx14(y, m)
                    except Exception as e:
                        print(f"✗ DTX14 の雛形: {e}")
                        errs.append(f"DTX14 の雛形: {e}")
            if errs:
                raise SystemExit("\n".join(errs))
            return None

        self.run_bg(job, self.b1, after=lambda: (self.refresh_months(), self._month1_changed()))

    # ---- ② 提出状況
    def _build2(self):
        f = self.t2
        row = ttk.Frame(f)
        row.pack(anchor="w", pady=4)
        ttk.Label(row, text="対象の月").pack(side="left")
        self.cb2 = ttk.Combobox(row, width=30, state="readonly")
        self.cb2.pack(side="left", padx=6)
        self._bind_month(self.cb2)
        ttk.Label(f, text="その月の全科の予定を読んで、次のことを下の「結果・ログ」に出します。\n"
                          "  ・まだ提出していない科\n"
                          "  ・「当直のみ」（日中は本院で勤務しない）と申告された医師\n"
                          "  ・名簿外の医師（Dr.JOY に登録してもらう必要がある人）\n"
                          "  ・重複や、取り込めなかった行（あれば）\n"
                          "同じ内容を 集約\\集約レポート.txt にも保存します。何度押してもかまいません（予定は変わりません）。",
                  foreground="#444", justify="left").pack(anchor="w", pady=(8, 4))
        self.b2 = ttk.Button(f, text="提出状況を確かめる", command=self.do2)
        self.b2.pack(anchor="e", pady=(8, 0))
        self.refresh_months()

    def do2(self):
        ym = self._cb_ym(self.cb2)
        if not ym:
            messagebox.showerror("入力", "対象の月を選んでください")
            return
        y, m = ym
        self.clear_log()

        def job():
            import aggregate as AG
            md = month_dir(y, m)
            if not os.path.isfile(master_json(y, m)):
                raise SystemExit(f"全科の予定（マスター）がありません: {master_json(y, m)}")
            out = os.path.join(md, "集約")
            os.makedirs(out, exist_ok=True)
            AG.aggregate(y, m, os.path.join(md, "提出"), out, C.newest_staffs(), report_only=True)
            return None                                 # フォルダは開かない（結果はログ欄に出る）

        self.run_bg(job, self.b2)

    # ---- ③ 確定
    def _build3(self):
        f = self.t3
        row = ttk.Frame(f)
        row.pack(anchor="w", pady=4)
        ttk.Label(row, text="対象の月").pack(side="left")
        self.cb3 = ttk.Combobox(row, width=30, state="readonly")
        self.cb3.pack(side="left", padx=6)
        self._bind_month(self.cb3)
        ttk.Label(f, text="病院長の決裁に出す当直勤務表を印刷する、その直前に押してください。\n"
                          "押すと、各科がそのとき出している予定が「元の割当」として焼き付きます。\n"
                          "このあと各科がアプリで変えた分は「交代申請」になり、当直交代申請.xlsx で「承認」したものだけが\n"
                          "毎日15時に当直勤務表（DTX14）へ反映されます。\n"
                          "押す前に「② 提出状況を確かめる」で、まだ提出していない科が無いか確かめてください。\n"
                          "「確定を解除」で元の状態に戻せるのは、交代の承認がまだ1件も無いときだけです。",
                  foreground="#444", justify="left").pack(anchor="w", pady=(8, 4))
        row2 = ttk.Frame(f)
        row2.pack(anchor="e", pady=(8, 0))
        self.b3u = ttk.Button(row2, text="確定を解除", command=lambda: self.do3(False))
        self.b3u.pack(side="left", padx=6)
        self.b3 = ttk.Button(row2, text="この月を確定", command=lambda: self.do3(True))
        self.b3.pack(side="left")
        self.refresh_months()

    def ask_big(self, title, big, body, ok_text, cancel_text="やめる"):
        """大きな字の1行（月など）と説明のついた「はい／やめる」の窓。既定のボタンは「やめる」。戻り: はいを押したか"""
        top = tk.Toplevel(self)
        top.title(title)
        top.transient(self)
        top.resizable(False, False)
        res = {"ok": False}
        fr = ttk.Frame(top, padding=18)
        fr.pack(fill="both", expand=True)
        tk.Label(fr, text=big, font=(FONT, 20, "bold"), fg="#b71c1c", justify="left").pack(anchor="w", pady=(0, 12))
        tk.Label(fr, text=body, font=(FONT, FONT_SIZE), justify="left", wraplength=int(760 * max(
            1.0, self.winfo_fpixels("1i") / 96.0))).pack(anchor="w")
        bf = ttk.Frame(fr)
        bf.pack(anchor="e", pady=(16, 0))

        def done(v):
            res["ok"] = v
            top.destroy()
        b_no = ttk.Button(bf, text=cancel_text, command=lambda: done(False))
        b_ok = ttk.Button(bf, text=ok_text, command=lambda: done(True))
        b_ok.pack(side="left", padx=(0, 12))
        b_no.pack(side="left")
        top.protocol("WM_DELETE_WINDOW", lambda: done(False))
        top.bind("<Escape>", lambda e: done(False))
        b_no.focus_set()
        top.grab_set()
        self.wait_window(top)
        return res["ok"]

    def do3(self, closed):
        ym = self._cb_ym(self.cb3)
        if not ym:
            messagebox.showerror("入力", "対象の月を選んでください")
            return
        y, m = ym
        name = f"R{y - 2018}年{m}月分（{y}年{m}月）" + ("（動作確認用テストフォルダ）" if (y, m) in test_dirs() else "")
        p = master_json(y, m)
        if not os.path.isfile(p):
            messagebox.showerror("確定", f"マスターがありません:\n{p}\n先に①で月フォルダを作ってください")
            return
        try:
            conf, cur, probs = month_confirmed(y, m)       # 確定の正本で決める（20260926）
        except MST.TruthError as e:
            if closed:
                messagebox.showerror("確定", f"確定の記録（確定.json）が読めません。管理者に知らせてください:\n{e}")
                return
            try:                                         # 解除は通す（確定.json を作り直す直し道。master_store.unconfirm）
                cur = MST.load(p, y, m)
            except Exception as e2:
                messagebox.showerror("確定", f"全科の予定（マスター）が読めません:\n{e2}")
                return
            conf = {"at": "確定.json が読めない"}
        except Exception as e:
            messagebox.showerror("確定", f"全科の予定（マスター）が読めません:\n{e}")
            return
        mconf = cur.get("confirmed")
        if closed:
            if conf:
                messagebox.showwarning("確定", f"{name}は、もう確定しています（{conf.get('at') or ''}）。")
                return
            if mconf:
                messagebox.showwarning("確定", f"{name}は、全科の予定に確定の印（{mconf.get('at') or ''}）が残っていますが、"
                                               "確定の記録（確定.json）は未確定です。\n"
                                               "先に「確定を解除」を押してから、もう一度「この月を確定」を押してください。")
                return
            miss = missing_owners(cur)
            dl = month_deadline(y, m)
            body = ("・各科がいま出している予定を「元の割当」として焼き付けます。\n"
                    "・このあと各科がアプリで変えた分は「交代申請」になり、当直交代申請.xlsx で承認したものだけが"
                    "当直勤務表に反映されます。\n"
                    "・病院長の決裁に出す当直勤務表を印刷する直前に押してください。\n")
            if dl and dl > dt.datetime.now():
                body += f"\n⚠ 締切（{dl.month}月{dl.day}日 {dl:%H:%M}）は、まだ来ていません。\n"
            body += ("\nまだ提出が無い科: " + (f"{len(miss)}（{'、'.join(miss)}）" if miss else "ありません（全科そろっています）"))
            if not self.ask_big("確定の確認", f"{name}を確定します", body, "確定する"):
                return
            if miss and not self.ask_big(
                    "もう一度確認",
                    f"まだ提出が無い科が {len(miss)} あります",
                    f"{'、'.join(miss)}\n\n"
                    "このまま確定すると、この科の「元の割当」は空のままになり、あとから出した予定はすべて"
                    "「交代申請」として扱われます（1件ずつ承認が要ります）。\n\n"
                    f"本当に{name}を確定しますか？",
                    "それでも確定する"):
                return
        else:
            if not conf and not mconf:
                messagebox.showinfo("確定を解除", f"{name}は、まだ確定していません（解除するものがありません）。")
                return
            if not self.ask_big(
                    "確定の解除の確認", f"{name}の確定を解除します",
                    "解除すると:\n"
                    "・確定する前と同じ扱いに戻ります。各科がアプリで直した分は、承認なしでそのまま毎日15時に"
                    "当直勤務表（DTX14）へ反映されます。\n"
                    "・当直交代申請.xlsx にすでに載っている申請は消えません。\n\n"
                    "もう一度「この月を確定」を押したとき:\n"
                    "・交代の承認がまだ1件も無ければ、そのときの予定で「元の割当」を焼き直します"
                    "（違う月を確定してしまったときは、これで元に戻せます）。\n"
                    "・承認した交代が1件でもあれば、元の割当は焼き直されず、承認した交代もそのまま残ります。"
                    "この場合は解除しても元どおりにはならないので、先に管理者に相談してください。",
                    "解除する"):
                return
        self.clear_log()
        who = os.environ.get("USERNAME", "")
        try:
            mst = (MST.confirm if closed else MST.unconfirm)(p, by=who, year=y, month=m)
        except MST.AlreadyConfirmed as e:      # 二重押し（焼き直すと元の割当と交代の記録が失われる）
            messagebox.showwarning("確定", str(e))
            self.logw(f"中止: {e}\n")
            return
        except Exception as e:
            messagebox.showerror("確定", str(e))
            self.logw(f"エラー: {e}\n")
            return
        n = sum(1 for e in (mst.get("depts") or {}).values() if e.get("duties"))
        self.logw((f"{name}を確定しました: " if closed else f"{name}の確定を解除しました: ") + p + "\n")
        if closed:
            self.logw(f"  元の割当を焼いた科: {n}／{len(C.owners_of(C.load_frames()))}（rev {mst.get('rev')}）\n")
        self.refresh_months()


    # ---- ④ 手入力の部署を取り込む（20260925 管理者: 旧「DTX14 に書き出す」は15時の反映がするので置き換えた）
    def _build4(self):
        f = self.t4
        row = ttk.Frame(f)
        row.pack(anchor="w", pady=4, fill="x")
        ttk.Label(row, text="対象の月").pack(side="left")
        self.cb4 = ttk.Combobox(row, width=12, state="readonly")
        self.cb4.pack(side="left", padx=6)
        self.cb4.bind("<<ComboboxSelected>>", lambda e: self._month4_changed())
        self.tg4 = ttk.Label(row, text="", foreground="#666")
        self.tg4.pack(side="left", padx=6)
        # 20261009: 事務担当の PC で月が出なかった（SharePoint の同期の場所が違う）→ 自動で探し、だめなら人が選んで覚える
        ttk.Button(row, text="当直勤務表_正本 の場所を選ぶ…", command=self.pick_root4).pack(side="right")
        ttk.Label(f, text="看護部・薬剤部・検査技師・放射線技師・管理 など、アプリで入力しない部署は、今までどおり DTX14 の形の Excel で届きます。\n"
                          "「ファイルを選んで取り込む…」でファイルを選ぶと、そのまま取り込みに進みます（最後に「よろしいですか？」と聞きます）。\n"
                          "その部署の行だけを、その月の当直勤務表（SharePoint の正本 DTX14）へ写します。\n"
                          "・1つのファイルに、いまの当直勤務表と違う部署が2つ以上あるときだけ、「どの部署を取り込むか」を選ぶ窓が出ます。\n"
                          "　届いたファイルに、ほかの部署の古い行が入っていることがあるためです。送ってきた部署だけを選んでください。\n"
                          "・各科（アプリで入力する科）と管理当直の行には触りません。\n"
                          "・行は A列の部署名と C列の見出し（昼・夜・深夜・日勤 など）で合わせます。10月までの形でも11月からの形でもかまいません。\n"
                          "・違う月の表は取り込みません。まるごと空の行は写しません（正本のまま）。\n"
                          "・書く前に正本を _上書き履歴 フォルダに写します。正本を Excel で開いている人がいるときは書きません。\n"
                          "・14時55分〜15時15分は取り込めません（15時の自動処理が同じ正本に書くため）。",
                  foreground="#444", justify="left").pack(anchor="w", pady=(6, 4))
        lf = ttk.Frame(f)
        lf.pack(fill="x", pady=(4, 0))
        self.lb4 = tk.Listbox(lf, height=4, selectmode="extended", font=(FONT, FONT_SIZE))
        self.lb4.pack(side="left", fill="x", expand=True)
        bf = ttk.Frame(lf)
        bf.pack(side="left", fill="y", padx=(6, 0))
        ttk.Button(bf, text="ファイルを選んで取り込む…", command=self.pick4).pack(fill="x")
        ttk.Button(bf, text="選んだものを外す", command=self.drop4).pack(fill="x", pady=(4, 0))
        row3 = ttk.Frame(f)
        row3.pack(anchor="e", pady=(8, 0))
        self.b4c = ttk.Button(row3, text="確かめる（書かない）", command=lambda: self.do4(False))
        self.b4c.pack(side="left", padx=6)
        self.b4 = ttk.Button(row3, text="取り込む", command=lambda: self.do4(True))
        self.b4.pack(side="left")
        self.files4 = []
        self._plan4 = None
        self.refresh4()

    def refresh4(self):
        ms = MI.list_months()
        vals = [f"R{y - 2018}.{m}月" for y, m in ms]
        self._months4 = dict(zip(vals, ms))
        cur = self.cb4.get()
        self.cb4["values"] = vals
        if cur not in vals:
            try:                                        # ②③と同じ既定（20260926。以前は①と同じ翌月・翌々月）
                dflt = default_month([(y, m, is_confirmed(y, m)) for y, m in ms])
            except Exception:
                dflt = None
            pick = f"R{dflt[0] - 2018}.{dflt[1]}月" if dflt else ""
            self.cb4.set(pick if pick in vals else (vals[-1] if vals else ""))
        self._month4_changed()

    def _month4_changed(self):
        ym = self._months4.get(self.cb4.get())
        if not ym:
            many = len(MI._SEARCH.get("hits") or []) > 1
            self.tg4.configure(text=("「当直勤務表_正本」がこの PC に2か所あります → 右の「場所を選ぶ…」で、SharePoint と同期している方を選んでください"
                                     if many else
                                     "SharePoint の同期フォルダ（当直勤務表_正本）がこの PC に見つかりません → 右の「場所を選ぶ…」で選んでください")
                               if not MI.dtx_root() else "", foreground="#c62828")
            return
        try:
            off = MI.official_path(*ym)
            # 20261009 レビュー: どのフォルダの正本に書くのか・今日の15時の版かを、事務担当が見られるように
            when = dt.datetime.fromtimestamp(os.path.getmtime(off)).strftime("%m/%d %H:%M")
            self.tg4.configure(text=f"正本: {os.path.basename(off)}（{when} 更新）\n場所: {os.path.dirname(os.path.dirname(off))}",
                               foreground="#666")
        except MI.ManualImportError as e:
            self.tg4.configure(text=str(e), foreground="#c62828")

    def pick_root4(self):
        d = filedialog.askdirectory(title="SharePoint の「当直勤務表_正本」フォルダを選んでください（エクスプローラーで同期している場所）")
        if not d:
            return
        d = os.path.normpath(d)
        if not MI._looks_like_root(d):
            messagebox.showwarning("場所を選ぶ", "このフォルダには「R8.11月 当直勤務表」のような月のフォルダがありません。\n"
                                              "「当直勤務表_正本」そのものを選んでください。\n\n" + d)
            return
        MI.remember_root(d)
        MI._SEARCH.clear()
        use_my_dtx_root()
        self.logw(f"④ 当直勤務表_正本 の場所: {d}（次からもここを使います）\n")
        self.refresh4()

    def pick4(self):
        """ファイルを選んだら、そのまま取り込みに進む（20261013 事務担当: 選んでからもう一度「取り込む」を押すのが分かりにくい）。
        選び直すと前の一覧は入れ替わる。複数のファイルは、ファイルを選ぶ窓でまとめて選ぶ。"""
        ps = filedialog.askopenfilenames(title="各部署から届いた DTX14 の形の Excel を選ぶ（いくつでも）",
                                         filetypes=[("Excel", "*.xlsm *.xlsx"), ("すべて", "*.*")])
        if not ps:
            return
        self.files4 = []
        self.lb4.delete(0, "end")
        for p in ps:
            p = os.path.normpath(p)
            if p not in self.files4:
                self.files4.append(p)
                self.lb4.insert("end", p)
        self.do4(True)

    def drop4(self):
        for i in sorted(self.lb4.curselection(), reverse=True):
            self.lb4.delete(i)
            del self.files4[i]

    def do4(self, write):
        ym = self._months4.get(self.cb4.get())
        if not ym:
            messagebox.showerror("入力", "対象の月を選んでください")
            return
        if not self.files4:
            messagebox.showerror("入力", "各部署から届いたファイルを選んでください")
            return
        if write and MI.in_quiet_window():
            messagebox.showwarning("取り込めない時間です", MI.QUIET_MESSAGE)
            return
        y, m = ym
        files = list(self.files4)
        self._plan4 = None
        self.clear_log()
        self._plan_bg(y, m, files, None, self.b4 if write else self.b4c,
                      after=self._choose_then_confirm4 if write else None)

    def _plan_bg(self, y, m, files, choose, btn, after=None):
        def plan_job():                                 # 読むだけ（正本には書かない）
            try:
                target = MI.official_path(y, m)
                if MI.locked(target):
                    print(f"⚠ 正本をいま Excel で開いている人がいます（閉じてもらうまで書けません）: {target}\n")
                plan = MI.plan_import(files, target, year=y, month=m, choose=choose)
            except MI.ManualImportError as e:
                raise SystemExit(str(e))
            print(MI.plan_text(plan))
            plan["_ym"], plan["_srcs"] = (y, m), files
            self._plan4 = plan
            return None

        self.run_bg(plan_job, btn, after=after)

    def _choose_then_confirm4(self):
        """1回目の読み（既定の選び方）のあと、ファイルごとに取り込む部署を選んでもらい、その選び方で計画を作り直す。"""
        plan = self._plan4
        if not plan:
            return
        if not any(f.get("depts") for f in plan["files"]):
            messagebox.showinfo("取り込み", "取り込める部署の行がありません（結果・ログを見てください）")
            return
        # 20261013 事務担当: 迷う余地が無いとき（どのファイルも、いまの当直勤務表と違う部署が1つ以下で、
        # 2つのファイルが同じ部署を取り合っていない）は、部署を選ぶ窓を出さずに最後の確認へ（確認にファイルと部署を出す）
        fls = [fl for fl in plan["files"] if fl.get("depts")]
        if (all(sum(1 for d in fl["depts"] if d["chosen"]) <= 1 for fl in fls)
                and not any(d["shared"] for fl in fls for d in fl["depts"])
                and any(d["chosen"] for fl in fls for d in fl["depts"])):
            self._confirm4()
            return
        choose = self.choose4(plan)
        if choose is None:
            self.logw("\n取り込みをやめました（正本は変わっていません）\n")
            return
        y, m = plan["_ym"]
        self._plan4 = None
        self.clear_log()
        self._plan_bg(y, m, plan["_srcs"], choose, self.b4, after=self._confirm4)

    def choose4(self, plan):
        """ファイルごとに、取り込む部署にチェックを入れてもらう窓。戻り: {ファイルのパス: {部署のキー}}（やめたら None）。"""
        top = tk.Toplevel(self)
        top.title("取り込む部署を選ぶ")
        top.transient(self)
        f = max(1.0, self.winfo_fpixels("1i") / 96.0)
        top.geometry(f"{min(int(900 * f), int(self.winfo_screenwidth() * 0.9))}x"
                     f"{min(int(620 * f), int(self.winfo_screenheight() * 0.85))}")
        tk.Label(top, text=f"R{plan['year'] - 2018}年{plan['month']}月分の当直勤務表に取り込む部署を選んでください",
                 font=(FONT, 16, "bold"), justify="left").pack(anchor="w", padx=14, pady=(12, 4))
        tk.Label(top, text="ファイルごとに、取り込む部署にチェックを入れます。\n"
                           "最初にチェックが入っているのは、そのファイルに書いてあって、いまの当直勤務表と違う部署です。\n"
                           "届いたファイルに ほかの部署の古い行が入っていることがあるので、送ってきた部署だけにしてください。\n"
                           "同じ部署を2つのファイルで選ぶことはできません。",
                 font=(FONT, FONT_SIZE), fg="#444", justify="left").pack(anchor="w", padx=14)
        body = ttk.Frame(top)
        body.pack(fill="both", expand=True, padx=14, pady=8)
        cv = tk.Canvas(body, highlightthickness=0)
        sb = ttk.Scrollbar(body, command=cv.yview)
        inner = ttk.Frame(cv)
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=inner, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-e.delta / 120), "units"))
        vars_ = []
        for fl in plan["files"]:
            lf = ttk.LabelFrame(inner, text=fl["name"], padding=8)
            lf.pack(fill="x", pady=4)
            if not fl["ok"]:
                ttk.Label(lf, text=f"取り込みません: {fl['reason']}", foreground="#c62828").pack(anchor="w")
                continue
            if not fl.get("depts"):
                ttk.Label(lf, text="取り込める部署の行がありません" + (f"（{fl['reason']}）" if fl["reason"] else ""),
                          foreground="#666").pack(anchor="w")
                continue
            for d in fl["depts"]:
                labels = list(dict.fromkeys(d["labels"]))
                note = "いまの当直勤務表と違う" if d["differs"] else "いまの当直勤務表と同じ（取り込んでも変わりません）"
                if d["shared"]:
                    note += f"／ほかのファイル（{'、'.join(d['shared'])}）にもあります"
                v = tk.BooleanVar(value=d["chosen"])
                ttk.Checkbutton(lf, text=f"{d['dept']}（{'・'.join(labels)}）… {note}", variable=v).pack(anchor="w")
                vars_.append((fl["path"], d["key"], d["dept"], v))
        res = {"v": None}

        def ok():
            pick, seen, dup = {}, {}, []
            for path, key, dept, v in vars_:
                pick.setdefault(path, set())
                if v.get():
                    pick[path].add(key)
                    if key in seen and seen[key] != path and dept not in dup:
                        dup.append(dept)
                    seen.setdefault(key, path)
            if dup:
                messagebox.showerror("選び直してください",
                                     f"同じ部署（{'、'.join(dup)}）を2つ以上のファイルで選んでいます。\n"
                                     "どれか1つのファイルだけにチェックを入れてください。", parent=top)
                return
            if not any(pick.values()):
                messagebox.showerror("選び直してください", "取り込む部署が1つも選ばれていません。", parent=top)
                return
            res["v"] = pick
            top.destroy()
        bf = ttk.Frame(top)
        bf.pack(anchor="e", padx=14, pady=(0, 12))
        ttk.Button(bf, text="選んだ部署を取り込む", command=ok).pack(side="left", padx=(0, 12))
        ttk.Button(bf, text="やめる", command=top.destroy).pack(side="left")
        top.bind("<Escape>", lambda e: top.destroy())
        top.grab_set()
        self.wait_window(top)
        try:
            cv.unbind_all("<MouseWheel>")
        except Exception:
            pass
        return res["v"]

    def _confirm4(self):
        plan = self._plan4
        if not plan:
            return
        if plan.get("conflicts"):
            messagebox.showerror("取り込み", "同じ行を2つ以上のファイルで選んでいます（結果・ログを見てください）。\n"
                                            "もう一度「取り込む」を押して選び直してください。")
            return
        n = len(plan["writes"])
        if not n:
            messagebox.showinfo("取り込み", "写す行がありません（結果・ログを見てください）")
            return
        if MI.in_quiet_window():
            messagebox.showwarning("取り込めない時間です", MI.QUIET_MESSAGE)
            return
        depts = list(dict.fromkeys(w["dept"] for w in plan["writes"]))
        bad = sum(len(f["unmatched"]) for f in plan["files"]) + sum(1 for f in plan["files"] if not f["ok"])
        per_file = []
        for fl in plan["files"]:
            ds = [f"{d['dept']}（{'・'.join(dict.fromkeys(d['labels']))}）" for d in fl.get("depts") or [] if d["chosen"]]
            if ds:
                per_file.append(f"・{fl['name']}\n　→ {'、'.join(ds)}")
        msg = (f"R{plan['year'] - 2018}年{plan['month']}月分の当直勤務表（正本 DTX14）に、\n"
               + ("\n".join(per_file) + "\n" if per_file else "")
               + f"{'、'.join(depts)} の {n} 行を写します。\n"
               + (f"合わなかった行・取り込まないファイルが {bad} 件あります（結果・ログを見てください）。\n" if bad else "")
               + "よろしいですか？")
        if not messagebox.askyesno("確認", msg):
            self.logw("\n取り込みをやめました（正本は変わっていません）\n")
            return
        target = plan["target"]

        def apply_job():
            print("")
            try:
                MI.apply_import(plan, target)
            except MI.ManualImportError as e:
                raise SystemExit(str(e))
            print("取り込みが終わりました。")
            return None

        self.run_bg(apply_job, self.b4)


SELFCHECK_MODULES = (
    "contract", "master_store", "month_setup", "make_month", "bucho_roster", "aggregate",
    "new_month", "dtx14_daily_sync", "dtx14_split_rows", "dtx14_manual_import", "dtx14_fill", "dtx14_extract",
    "kotai_record", "set_acl", "duty_sheet", "sheet_style", "make_staff_csv",
    "openpyxl", "win32com.client", "pythoncom", "pywintypes", "win32timezone",
)


def selfcheck(out_path, com=False, staffs=None):
    """exe のビルドの確かめ（画面は出さない）。使うモジュールを全部読み込み、同梱の data を読む。
    com=True なら Excel（新しい空のブック・保存しない）で日付を1つ読む＝COM の日付の変換（win32timezone）が動くか。
    staffs を渡せば make_month.build_config も通す（書かない）。結果を out_path に書く。戻り: 終了コード"""
    import importlib
    import traceback
    lines, ok = [f"admin_gui selfcheck frozen={getattr(sys, 'frozen', False)} exe={sys.executable}"], True
    for name in SELFCHECK_MODULES:
        try:
            importlib.import_module(name)
            lines.append(f"OK import {name}")
        except Exception as e:
            ok = False
            lines.append(f"NG import {name}: {type(e).__name__}: {e}")
    try:
        fr = C.load_frames()
        hol = C.load_holidays()
        lines.append(f"OK data: frames={len(fr)} holidays={len(hol)} owners={len(C.owners_of(fr))}"
                     f" deadline(2026-11)={NM.deadline_for(2026, 11)} DATA={C.DATA}")
    except Exception as e:
        ok = False
        lines.append(f"NG data: {type(e).__name__}: {e}")
    if staffs:
        try:
            import make_month as MM
            logs = []
            cfg = MM.build_config(2026, 11, "", staffs, None, log=logs.append)
            lines.append(f"OK make_month.build_config: days={len(cfg['days'])} roster={len(cfg['roster'])}"
                         f" fixed={ {k: len(v) for k, v in cfg['fixedMembers'].items()} }")
            lines += ["   " + s for s in logs]
        except Exception as e:
            ok = False
            lines.append(f"NG make_month.build_config: {type(e).__name__}: {e}\n{traceback.format_exc()}")
    if com:
        try:
            del sys.modules["win32timezone"]            # 上で読み込んだ分を忘れ、COM の日付の変換で読み直させる
        except KeyError:
            pass
        try:
            import win32com.client
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            try:
                wb = xl.Workbooks.Add()
                ws = wb.Worksheets(1)
                ws.Range("A1").Formula = "=DATE(2026,11,1)"
                v = ws.Range("A1").Value
                wb.Close(SaveChanges=False)
            finally:
                xl.Quit()
            good = hasattr(v, "year") and (v.year, v.month, v.day) == (2026, 11, 1)
            ok = ok and good
            lines.append(("OK" if good else "NG") + f" COM date read: {v!r}")
        except Exception as e:
            ok = False
            lines.append(f"NG COM date read: {type(e).__name__}: {e}")
    lines.append("RESULT " + ("OK" if ok else "NG"))
    with open(out_path, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write("\n".join(lines) + "\n")
    return 0 if ok else 1


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--selfcheck":     # 隠しの確かめ（exe のビルド用。画面は出さない）
        args = sys.argv[3:]
        staffs = next((a for a in args if not a.startswith("--")), None)
        sys.exit(selfcheck(sys.argv[2], com="--com" in args, staffs=staffs))
    App().mainloop()


if __name__ == "__main__":
    main()
