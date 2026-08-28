"""
관리계획서(CP) -> 셋업 & 일일 성형 조건 기록지 생성기

정책 (사용자 확정 2026-08-28):
    - CP의 '관리기록방법'(AJ열) 값이 '셋업 & 일일 성형 조건 기록지'인 행만 위에서부터 순서대로 뽑는다.
      (라벨을 코드에 넣어두지 않고, CP가 바뀌면 그대로 반영되도록 CP에서 직접 읽는다.)
    - No 열      : 뽑힌 행을 위에서부터 1, 2, 3 ...
    - 공정단계 열 : CP 공정명(B열). 공정그룹 첫 행에만 값이 있어(병합) 직전 값을 이어받는다.
    - 공정특성 열 : CP 공정특성(L열). 검토표시본 주석(🔍/⚠️/❌, [REVIEW], [MISSING])은 잘라낸다.
    - S1/R · S2  : CP 주기 표시(AC=S1, AD=S2)에 'X'가 있으면 그대로 옮긴다.
    - 규격치/공차/단위 : CP 규격하한/Norminal/규격상한/단위(T/U/W/X)로부터 계산.
    - 표의 헤더 행 / 데이터 시작 행 / 각 열 위치는 템플릿에서 자동으로 찾는다
      (헤더 행에서 'No/S1·R/S2/공정단계/공정특성/규격치/공차/단위' 라벨 스캔).
      그래서 템플릿 상단 구성이 바뀌어도 데이터가 밀리지 않는다.
    - 표(테두리) 길이는 뽑힌 항목 수에 맞춘다. 템플릿이 미리 그려둔 빈 행이 더 많으면 지우고,
      항목이 더 많으면 첫 데이터행 서식을 복제해 늘린다.
    - 날짜별 성형 기록 칸은 작업자 수기 입력용이라 비워 둔다.
"""
import re
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Border, PatternFill

from cp_loader import load_cp_workbook

CP_HEADER = {"제품번호": "A5", "대상설비": "M3"}

CP_DATA_START_ROW = 11
CP_COL = {
    "공정단계": "B",
    "공정특성": "L",
    "규격하한": "T",
    "Norminal": "U",
    "규격상한": "W",
    "단위": "X",
    "S1": "AC",
    "S2": "AD",
    "관리기록방법": "AJ",
}

# 관리기록방법(공백 제거)에 이 문자열이 들어있으면 성형조건기록지 항목으로 본다.
TARGET_METHOD_KEY = "성형조건"

ANNOT_MARKERS = ("🔍", "⚠️", "❌", "✅", "[REVIEW", "[MISSING")
BLANK_TOKENS = ("", "-", "–", "—", "－", "etc")

# 헤더 행에서 찾을 열 라벨(공백 제거 후 비교) / 못 찾을 때 쓰는 기본 열 번호
LAYOUT_LABELS = {
    "no": "No",
    "s1": "S1/R",
    "s2": "S2",
    "step": "공정단계",
    "char": "공정특성",
    "spec": "규격치",
    "tol": "공차",
    "unit": "단위",
}
DEFAULT_LAYOUT = {"no": 2, "s1": 3, "s2": 4, "step": 5, "char": 6, "spec": 7, "tol": 8, "unit": 9}
DEFAULT_DATA_START = 4


def read_cp_header(ws):
    return {key: ws[coord].value for key, coord in CP_HEADER.items()}


def _norm(v):
    return re.sub(r"\s+", "", "" if v is None else str(v))


def clean_text(v):
    """검토표시본 주석(🔍/⚠️/❌ 등)을 잘라내고 공백/개행을 정리한다."""
    if v is None:
        return ""
    s = str(v).replace("\r", "")
    cut = len(s)
    for marker in ANNOT_MARKERS:
        i = s.find(marker)
        if i != -1 and i < cut:
            cut = i
    return " ".join(s[:cut].split())


def is_blank(v):
    return v is None or clean_text(v).strip().lower() in BLANK_TOKENS


def _to_float(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def _num(x):
    return int(x) if float(x).is_integer() else round(x, 4)


def _numbers(s):
    return [float(x) for x in re.findall(r"-?\d+(?:\.\d+)?", s or "")]


def _pair_tol(lo_s, nom_s, up_s):
    """'Front : 47/ Rear : 43' 처럼 값이 여러 개인 규격에서 공차를 계산한다."""
    nom, lo, up = _numbers(nom_s), _numbers(lo_s), _numbers(up_s)
    if len(nom) < 2:
        return None
    ups = [round(u - n, 4) for n, u in zip(nom, up)] if len(up) == len(nom) else []
    los = [round(n - l, 4) for l, n in zip(lo, nom)] if len(lo) == len(nom) else []
    diffs = ups + los
    if not diffs:
        return None
    if all(abs(d - diffs[0]) < 1e-9 for d in diffs) and diffs[0] >= 0:
        return f"± {diffs[0]:g}"
    if ups and los and all(abs(d - ups[0]) < 1e-9 for d in ups) and all(abs(d - los[0]) < 1e-9 for d in los):
        return f"+ {ups[0]:g} / - {los[0]:g}"
    return None


def compute_spec(lower, nominal, upper, unit):
    """CP 규격하한/Norminal/규격상한/단위 -> (규격치, 공차, 단위)."""
    lo_s, nom_s, up_s = clean_text(lower), clean_text(nominal), clean_text(upper)
    unit_s = "" if is_blank(unit) else clean_text(unit)

    lo_f, nom_f, up_f = _to_float(lo_s), _to_float(nom_s), _to_float(up_s)

    # 하한 / 상한 둘 다 숫자 -> 규격치 = Nominal(없으면 중앙값), 공차 = ±(상한-규격치)
    if lo_f is not None and up_f is not None:
        center = nom_f if nom_f is not None else (lo_f + up_f) / 2
        up_tol = round(up_f - center, 4)
        lo_tol = round(center - lo_f, 4)
        if abs(up_tol - lo_tol) < 1e-9:
            tol = f"± {up_tol:g}"
        else:
            tol = f"+ {up_tol:g} / - {lo_tol:g}"
        return _num(center), tol, unit_s or None

    # Nominal 만 숫자
    if nom_f is not None:
        return _num(nom_f), None, unit_s or None

    # 한쪽 경계만 숫자
    if lo_f is not None:
        return _num(lo_f), "이상", unit_s or None
    if up_f is not None:
        return _num(up_f), "이하", unit_s or None

    # 텍스트 규격 (Front/Rear 처럼 값이 여러 개면 공차 계산 시도)
    if not is_blank(nom_s):
        return nom_s, _pair_tol(lo_s, nom_s, up_s), unit_s or None
    if not is_blank(lo_s) and not is_blank(up_s):
        return f"{lo_s} ~ {up_s}", None, unit_s or None
    return None, None, None


def extract_rows(cp_ws):
    """관리기록방법이 성형조건기록지인 CP 행을 위에서부터 순서대로 수집한다."""
    rows = []
    current_step = ""
    for r in range(CP_DATA_START_ROW, cp_ws.max_row + 1):
        step_val = cp_ws[f'{CP_COL["공정단계"]}{r}'].value
        if not is_blank(step_val):
            current_step = clean_text(step_val)

        method = _norm(clean_text(cp_ws[f'{CP_COL["관리기록방법"]}{r}'].value))
        if TARGET_METHOD_KEY not in method:
            continue

        char_raw = cp_ws[f'{CP_COL["공정특성"]}{r}'].value
        char = clean_text(char_raw)
        if is_blank(char):
            continue

        spec, tol, unit = compute_spec(
            cp_ws[f'{CP_COL["규격하한"]}{r}'].value,
            cp_ws[f'{CP_COL["Norminal"]}{r}'].value,
            cp_ws[f'{CP_COL["규격상한"]}{r}'].value,
            cp_ws[f'{CP_COL["단위"]}{r}'].value,
        )
        rows.append(
            {
                "공정단계": current_step,
                "공정특성": char,
                "규격치": spec,
                "공차": tol,
                "단위": unit,
                "s1": not is_blank(cp_ws[f'{CP_COL["S1"]}{r}'].value),
                "s2": not is_blank(cp_ws[f'{CP_COL["S2"]}{r}'].value),
                "검토필요": "[MISSING" in str(char_raw) or "[REVIEW" in str(char_raw) or spec is None,
            }
        )
    return rows


def _match_col(header_norms, label):
    """헤더 라벨을 공백 무시 + 접두어 허용으로 매칭 ('공정단계' <-> '공정단계명')."""
    key = _norm(label)
    if key in header_norms:
        return header_norms[key]
    for hk, col in header_norms.items():
        if hk.startswith(key) or key.startswith(hk):
            return col
    return None


def resolve_layout(ws):
    """헤더 행에서 열 위치와 데이터 시작 행을 찾는다. 못 찾으면 기본값 사용."""
    header_row = None
    found = {}
    for row in ws.iter_rows(min_row=1, max_row=15):
        norms = {_norm(c.value): c.column for c in row if c.value not in (None, "")}
        if _match_col(norms, "공정특성") and _match_col(norms, "규격치"):
            header_row, found = row[0].row, norms
            break

    layout = {key: (_match_col(found, label) or DEFAULT_LAYOUT[key]) for key, label in LAYOUT_LABELS.items()}

    if header_row is None:
        layout["data_start"] = DEFAULT_DATA_START
        return layout

    # 헤더가 병합으로 여러 행을 차지하면 그 아래부터
    base = header_row + 1
    for m in ws.merged_cells.ranges:
        if m.min_row <= header_row <= m.max_row and m.min_col <= layout["unit"]:
            base = max(base, m.max_row + 1)
    # 주석 행('(S1/S2/R ...)' 처럼 괄호로 시작하는 안내)은 건너뛴다
    for _ in range(5):
        texts = [
            str(ws.cell(row=base, column=col).value).strip()
            for col in range(layout["no"], layout["unit"] + 1)
            if ws.cell(row=base, column=col).value not in (None, "")
        ]
        if texts and all(t[:1] in "(`" for t in texts):
            base += 1
        else:
            break
    layout["data_start"] = base
    return layout


def _set(ws, col, row, value):
    cell = ws.cell(row=row, column=col)
    if type(cell).__name__ != "MergedCell":
        cell.value = value


def fill_rows(ws, rows, layout):
    for idx, row in enumerate(rows):
        r = layout["data_start"] + idx
        _set(ws, layout["no"], r, idx + 1)
        _set(ws, layout["step"], r, row["공정단계"])
        _set(ws, layout["char"], r, row["공정특성"])
        if row["s1"]:
            _set(ws, layout["s1"], r, "X")
        if row["s2"]:
            _set(ws, layout["s2"], r, "X")
        if row["규격치"] is not None:
            _set(ws, layout["spec"], r, row["규격치"])
        if row["공차"] is not None:
            _set(ws, layout["tol"], r, row["공차"])
        if row["단위"] is not None:
            _set(ws, layout["unit"], r, row["단위"])


def _has_border(cell):
    b = cell.border
    return any(getattr(b, s) and getattr(b, s).style for s in ("left", "right", "top", "bottom"))


def fit_table_rows(ws, layout, n_items):
    """표(테두리) 길이를 항목 수에 맞춘다. 남는 템플릿 행은 지우고, 모자라면 서식을 복제해 늘린다."""
    data_start = layout["data_start"]
    max_col = ws.max_column
    last_data_row = data_start + n_items - 1

    # 템플릿이 미리 테두리를 그려둔 마지막 행
    tpl_last = data_start - 1
    r = data_start
    while r <= ws.max_row and any(_has_border(ws.cell(row=r, column=c)) for c in range(1, max_col + 1)):
        tpl_last = r
        r += 1

    # 항목보다 템플릿 그리드가 길면 -> 남는 행의 값/서식 제거
    for rr in range(last_data_row + 1, tpl_last + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=rr, column=c)
            cell.value = None
            cell.border = Border()
            cell.fill = PatternFill()
        if rr in ws.row_dimensions:
            ws.row_dimensions[rr].height = None

    # 항목이 템플릿 그리드보다 많으면 -> 첫 데이터행 서식을 복제해 늘림
    for rr in range(tpl_last + 1, last_data_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=rr, column=c)._style = copy(ws.cell(row=data_start, column=c)._style)
        src_h = ws.row_dimensions[data_start].height if data_start in ws.row_dimensions else None
        if src_h is not None:
            ws.row_dimensions[rr].height = src_h


def generate(cp_path, template_path, output_path, sheet_name=None):
    cp_wb = load_cp_workbook(cp_path)
    cp_ws = cp_wb.active
    cp_header = read_cp_header(cp_ws)
    rows = extract_rows(cp_ws)

    wb = openpyxl.load_workbook(template_path, data_only=False)
    target_sheet = sheet_name or wb.sheetnames[-1]
    ws = wb[target_sheet]

    layout = resolve_layout(ws)
    fill_rows(ws, rows, layout)
    fit_table_rows(ws, layout, len(rows))

    wb.save(output_path)

    return {
        "제품번호": cp_header.get("제품번호"),
        "시트명": target_sheet,
        "데이터시작행": layout["data_start"],
        "항목수": len(rows),
        "CP에서_못찾은_항목": [row["공정특성"] for row in rows if row["검토필요"]],
    }


if __name__ == "__main__":
    cp_path, template_path, output_path = sys.argv[1], sys.argv[2], sys.argv[3]
    print(generate(cp_path, template_path, output_path))
