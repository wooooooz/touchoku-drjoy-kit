'use strict';
// rights.cjs — 貴科の医師リストを直せる人（20261031）
//   管理者 20261028: 貴科の医師リスト（Dr.JOY所属の申告）の × と区分の切り替えは、部門管理者だけができる。
//   部門管理者 = Dr.JOY の管理権限が「所属管理者」（その所属の秘書・部長）か「全体管理者」（管理者・事務担当・甲野など）。
//   一覧は毎朝3:30の名簿の更新が 名簿\roster_editors.json に置く（tools/roster_editors.py）。
//   交代で入る先生を足すことは誰でもできる（交代申請は本人が出す）。所属の申告にならない提出単位（集中治療室など）は制限しない。

// 区分（contract.work_of と同じ規則）: work があればそれ、無ければ in_house=false → part、それ以外 → full
function workOf(m) {
  const w = m && m.work;
  if (w === 'full' || w === 'part' || w === 'ingai') return w;
  return m && m.in_house === false ? 'part' : 'full';
}

// 集合当直で選んだ「Dr.JOY所属」（行の seg。20261031）。区分と同じく、変えられるのは部門管理者だけ
const segOf = (m) => String((m && m.seg) || '');

// 前の名簿 → 新しい名簿で、外した人と区分・「Dr.JOY所属」を変えた人（足した人は問わない）
function rosterDiff(before, after) {
  const now = new Map((after || []).map((m) => [String(m.login || '').toLowerCase(), m]));
  const removed = [];
  const changed = [];
  for (const m of before || []) {
    const k = String(m.login || '').toLowerCase();
    if (!k) continue;
    if (!now.has(k)) removed.push(m);
    else if (workOf(now.get(k)) !== workOf(m) || segOf(now.get(k)) !== segOf(m)) changed.push({ before: m, after: now.get(k) });
  }
  return { removed, changed };
}

// dept_map の1項目（所属パターンの配列、または {dept:[...], rank:[...]}）→ 所属パターン
function patternsOf(entry) {
  if (Array.isArray(entry)) return entry;
  return (entry && Array.isArray(entry.dept)) ? entry.dept : [];
}

const segs = (d) => String(d || '').split(/[/／、]/).map((s) => s.trim()).filter(Boolean);

// この人（editor = roster_editors.json の1人。null=管理者でない）が、その提出単位の部門管理者か
//   全体管理者・追加の作成者（owners に名指し）・所属管理者（所属が提出単位の dept_map に当たる）。20261028
function canEdit(owner, editor, deptMap) {
  if (!editor) return false;
  if (editor.level === '全体') return true;
  if ((editor.owners || []).includes(owner)) return true;
  if (editor.level !== '所属') return false;
  const pats = patternsOf((deptMap || {})[owner]);
  return (editor.depts || []).some((d) => segs(d).some((s) => pats.some((p) => p && s.includes(p))));
}

// 貴科の医師リストで医師を外す・区分を変えることができるか（決裁の前も後も）。
//   shozoku = config の shozoku（owners = 所属の申告になる提出単位）。所属の申告にならない提出単位は制限しない
function canEditRoster(owner, editor, shozoku, deptMap) {
  const owners = (shozoku && shozoku.owners) || [];
  if (!owners.includes(owner)) return true;
  return canEdit(owner, editor, deptMap);
}

// 決裁前の月の当直予定とメンバーの変更が、前の版から変わっているか（変わっていれば部門管理者だけ。20261028）
function entryChanged(before, ent) {
  const dk = (r) => [r.frame_id, r.date, r.band, String(r.staff_login || '').toLowerCase(), r.oncall ? 1 : 0].join('|');
  const rk = (m) => `${String(m.login || '').toLowerCase()}|${workOf(m)}|${segOf(m)}`;
  const same = (a, b, f) => JSON.stringify((a || []).map(f).sort()) === JSON.stringify((b || []).map(f).sort());
  const b = before || {};
  const d = Array.isArray(ent.duties) ? ent.duties : (b.duties || []);      // 送ってこなかった方は前のまま
  const r = Array.isArray(ent.roster) ? ent.roster : (b.roster || []);
  return !(same(b.duties, d, dk) && same(b.roster, r, rk));
}

// Windows のログイン名（user001 など）→ roster_editors.json の1人
function editorOf(editors, user) {
  const k = String(user || '').trim().toLowerCase();
  return (k && editors && editors[k]) || null;
}

module.exports = { workOf, rosterDiff, patternsOf, canEdit, canEditRoster, entryChanged, editorOf };
