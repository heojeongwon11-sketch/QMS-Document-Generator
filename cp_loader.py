"""
CP(관리계획서) 파일 로더 — .xls(구형) / .xlsx 둘 다 지원.

generate_is.py, generate_screw.py 등에서 CP 파일을 열 때
openpyxl.load_workbook(cp_path) 대신 load_cp_workbook(cp_path) 를 쓰면
확장자와 무관하게(실제 파일 시그니처로 판별) 동일한 방식으로 셀 값을 읽을 수 있다.

.xls 파일은 서식/수식 없이 '값'만 옮겨서 openpyxl Workbook 형태로 재구성한다.
(CP 파일은 어차피 값만 읽기 때문에 서식은 필요 없음)
"""
import io
import os

import openpyxl

try:
    import xlrd
except ImportError:  # xlrd 가 없으면 .xls 지원만 불가능, .xlsx 는 정상 동작
    xlrd = None

XLSX_SIGNATURE = b"PK\x03\x04"
XLS_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _read_bytes(source):
    """path(str) / file-like(BytesIO, streamlit UploadedFile) / bytes 모두 지원."""
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if isinstance(source, (str, os.PathLike)):
        with open(source, "rb") as f:
            return f.read()
    if hasattr(source, "read"):
        try:
            source.seek(0)
        except Exception:
            pass
        data = source.read()
        try:
            source.seek(0)
        except Exception:
            pass
        return data
    raise TypeError(f"지원하지 않는 입력 형식입니다: {type(source)}")


def _xls_bytes_to_openpyxl(data: bytes):
    if xlrd is None:
        raise RuntimeError(
            "이 파일은 구형 .xls 형식인데 xlrd 패키지가 설치되어 있지 않습니다. "
            "터미널에서 'pip install xlrd' 실행 후 다시 시도해 주세요."
        )
    book = xlrd.open_workbook(file_contents=data)
    sheet = book.sheet_by_index(0)

    wb = openpyxl.Workbook()
    ws = wb.active

    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            value = sheet.cell_value(r, c)
            if value == "":
                continue
            cell_type = sheet.cell_type(r, c)
            if cell_type == xlrd.XL_CELL_DATE:
                value = xlrd.xldate_as_datetime(value, book.datemode)
            elif cell_type == xlrd.XL_CELL_NUMBER:
                # 정수로 딱 떨어지면 int로 (엑셀에서 "135" 처럼 보이는 값이 135.0으로 안 읽히게)
                if isinstance(value, float) and value.is_integer():
                    value = int(value)
            ws.cell(row=r + 1, column=c + 1, value=value)

    return wb


def load_cp_workbook(source):
    """
    source: 파일 경로(str) / 파일 객체(BytesIO, streamlit UploadedFile) / bytes
    반환: openpyxl Workbook (기존 openpyxl.load_workbook 사용부와 동일하게 다룰 수 있음)
    """
    data = _read_bytes(source)

    if data[:4] == XLSX_SIGNATURE:
        try:
            return openpyxl.load_workbook(io.BytesIO(data), data_only=False)
        except Exception as e:
            raise ValueError(
                "CP 파일을 열 수 없습니다. 확장자는 .xlsx이지만 내용이 손상되었거나, "
                "실제로는 다른 형식(.xls를 이름만 바꾼 경우, 비밀번호 보호, 특수 OOXML 등)일 수 있습니다. "
                "엑셀에서 파일을 열어 '다른 이름으로 저장 → Excel 통합 문서(.xlsx)'로 다시 저장한 뒤 "
                f"업로드해 주세요. (원본 오류: {e})"
            ) from e
    if data[:8] == XLS_SIGNATURE:
        return _xls_bytes_to_openpyxl(data)

    raise ValueError(
        "CP 파일 형식을 인식할 수 없습니다. .xls 또는 .xlsx 파일인지 확인해 주세요. "
        f"(파일 시작 바이트: {data[:8]!r})"
    )
