# -*- coding: utf-8 -*-
"""
관로번호 검사 프로그램
- 엑셀의 '관로번호' 열과 동영상 파일명 시작 부분의 관로번호를 비교합니다.
- 원본 엑셀 및 동영상 파일은 읽기만 하며 수정/삭제하지 않습니다.
"""
import os
import re
import threading
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from collections import defaultdict

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    load_workbook = None

VIDEO_EXTENSIONS = {
    ".mp4", ".avi", ".mkv", ".mov", ".wmv", ".m4v", ".mpg", ".mpeg",
    ".3gp", ".ts", ".mts", ".m2ts", ".webm", ".vob", ".flv"
}

# 파일명 첫 토큰에서 영문/숫자 관로번호를 추출
# 예: NH-01.mp4 -> NH-01, ABC123 9M.mp4 -> ABC123, 202101900-1.mp4 -> 202101900-1
# 첫 띄어쓰기 뒤의 설명은 무시하고, 비교는 대소문자 구분 없이 수행
ID_PATTERN = re.compile(r"^([A-Za-z0-9]+(?:-[A-Za-z0-9]+)?)")
VALID_ID_PATTERN = re.compile(r"^[A-Za-z0-9]+(?:-[A-Za-z0-9]+)?$")

def canonical_id(value):
    """관로번호 비교용 정규형: 앞뒤 공백 제거 및 영문 대문자화."""
    return str(value).strip().upper()


def normalize_excel_id(value):
    """엑셀 셀 값을 관로번호 문자열로 정규화. 숫자형 셀도 처리."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return ""
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()
    text = str(value).strip()
    # 엑셀 숫자 값이 '202101900.0'처럼 들어온 경우 처리
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def extract_video_id(filename):
    stem = Path(filename).stem.strip()
    # 첫 띄어쓰기 앞부분만 ID로 취급. 이후 설명은 비교 대상에서 제외.
    token = stem.split(None, 1)[0] if stem else ""
    match = ID_PATTERN.match(token)
    return canonical_id(match.group(1)) if match else ""


def get_actual_max_row(excel_path, worksheet_xml_path):
    """
    Excel can report over a million rows because of formatting applied to an empty row.
    Scan worksheet XML and return the last row containing an actual value/formula.
    This avoids iterating through the inflated worksheet dimension.
    """
    max_data_row = 1
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(excel_path, "r") as archive:
        with archive.open(worksheet_xml_path) as stream:
            for event, elem in ET.iterparse(stream, events=("end",)):
                if elem.tag == ns + "row":
                    row_num = int(elem.attrib.get("r", "0") or 0)
                    has_value = False
                    for cell in elem.findall(ns + "c"):
                        # Cells containing only style information are not data.
                        if (cell.find(ns + "v") is not None or
                            cell.find(ns + "f") is not None or
                            cell.find(ns + "is") is not None):
                            has_value = True
                            break
                    if has_value:
                        max_data_row = max(max_data_row, row_num)
                    elem.clear()
    return max_data_row


def find_header_and_rows(excel_path):
    """Find the '관로번호' header, then read only the worksheet's real populated range."""
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    candidates = []
    try:
        # Determine the actual populated row range from XML rather than trusting
        # ws.max_row, which may be inflated by formatting-only rows.
        actual_max_rows = {}
        for ws in wb.worksheets:
            actual_max_rows[ws.title] = get_actual_max_row(excel_path, ws._worksheet_path)

        # The header is expected near the top of a survey sheet.
        for ws in wb.worksheets:
            max_row = min(actual_max_rows[ws.title], 200)
            max_col = min(ws.max_column or 100, 100)
            for row_idx, row in enumerate(
                ws.iter_rows(min_row=1, max_row=max_row, min_col=1,
                             max_col=max_col, values_only=True), start=1
            ):
                for col_idx, value in enumerate(row, start=1):
                    if value is not None and str(value).strip() == "관로번호":
                        candidates.append((ws.title, row_idx, col_idx))
                        break
                if candidates and candidates[-1][0] == ws.title:
                    break

        if not candidates:
            raise ValueError(
                "엑셀 파일의 상단 200행에서 '관로번호' 헤더를 찾지 못했습니다. "
                "시트와 헤더 이름을 확인해 주세요."
            )

        # Prefer a sheet explicitly marked '수정' when both original and edited
        # sheets contain the same header; otherwise use the earliest header.
        candidates.sort(key=lambda x: (
            0 if "수정" in x[0] else 1, x[1], x[2]
        ))
        sheet_name, header_row, col_idx = candidates[0]
        ws = wb[sheet_name]
        actual_max_row = actual_max_rows[sheet_name]
        records = []

        for row_idx, values in enumerate(
            ws.iter_rows(
                min_row=header_row + 1,
                max_row=actual_max_row,
                min_col=col_idx,
                max_col=col_idx,
                values_only=True,
            ),
            start=header_row + 1,
        ):
            raw = values[0] if values else None
            code = normalize_excel_id(raw)
            if not code:
                continue
            if code == "관로번호" or not VALID_ID_PATTERN.fullmatch(code):
                continue
            code = canonical_id(code)
            records.append({"sheet": sheet_name, "row": row_idx, "id": code})

        if not records:
            raise ValueError(
                "'관로번호' 열은 찾았지만 그 아래에서 관로번호 데이터를 찾지 못했습니다."
            )
        return records, sheet_name, header_row, col_idx
    finally:
        wb.close()


def scan_videos(folder, recursive=True):
    by_id = defaultdict(list)
    no_id = []
    root = Path(folder)
    iterator = root.rglob("*") if recursive else root.iterdir()
    for path in iterator:
        if not path.is_file() or path.suffix.lower() not in VIDEO_EXTENSIONS:
            continue
        code = extract_video_id(path.name)
        if code:
            by_id[code].append(str(path))
        else:
            no_id.append(str(path))
    for code in by_id:
        by_id[code].sort(key=str.casefold)
    no_id.sort(key=str.casefold)
    return by_id, no_id


def style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="D9E2F3")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col[:2000]:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, min(len(val), 65))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 55)


def write_report(output_path, excel_path, video_folder, records, by_id, no_id, sheet_name, header_row):
    wb = Workbook()
    summary = wb.active
    summary.title = "전체 검사 결과"
    summary.append(["엑셀 시트", "엑셀 행 번호", "관로번호", "동영상 개수", "검사 결과", "동영상 파일명(전체 경로)"])
    missing_rows = []
    multiple_rows = []
    matched_rows = []
    for rec in records:
        code = rec["id"]
        paths = by_id.get(code, [])
        count = len(paths)
        status = "영상 없음" if count == 0 else ("영상 여러 개" if count > 1 else "정상")
        filenames = "\n".join(paths)
        row = [rec["sheet"], rec["row"], code, count, status, filenames]
        summary.append(row)
        if count == 0:
            missing_rows.append(row)
        elif count > 1:
            multiple_rows.append(row)
        else:
            matched_rows.append(row)

    missing = wb.create_sheet("영상 없는 관로")
    missing.append(["엑셀 시트", "엑셀 행 번호", "관로번호", "동영상 개수", "검사 결과", "동영상 파일명(전체 경로)"])
    for row in missing_rows:
        missing.append(row)

    multiple = wb.create_sheet("영상 여러 개")
    multiple.append(["엑셀 시트", "엑셀 행 번호", "관로번호", "동영상 개수", "검사 결과", "동영상 파일명(전체 경로)"])
    for row in multiple_rows:
        multiple.append(row)

    # 실제 폴더에 있지만 엑셀의 관로번호에는 없는 영상 ID도 별도 표시
    excel_ids = {rec["id"] for rec in records}
    extra = wb.create_sheet("엑셀에 없는 영상")
    extra.append(["파일에서 추출한 관로번호", "파일 경로", "비고"])
    for code in sorted(by_id.keys()):
        if code not in excel_ids:
            for path in by_id[code]:
                extra.append([code, path, "동영상은 있으나 엑셀 관로번호 목록에 없음"])

    no_id_ws = wb.create_sheet("번호 추출 불가 영상")
    no_id_ws.append(["파일 경로", "비고"])
    for path in no_id:
        no_id_ws.append([path, "파일명 시작 부분에서 숫자 관로번호를 추출하지 못함"])

    info = wb.create_sheet("검사 정보")
    info.append(["항목", "내용"])
    info.append(["엑셀 파일", os.path.abspath(excel_path)])
    info.append(["동영상 폴더", os.path.abspath(video_folder)])
    info.append(["관로번호가 있는 시트", sheet_name])
    info.append(["관로번호 헤더 행", header_row])
    info.append(["엑셀 관로번호 행 수", len(records)])
    info.append(["영상 없음", len(missing_rows)])
    info.append(["영상 여러 개", len(multiple_rows)])
    info.append(["관로번호가 추출된 영상 파일 수", sum(len(v) for v in by_id.values())])
    info.append(["번호 추출 불가 영상 파일 수", len(no_id)])
    info.append(["번호 추출 규칙", "파일명 첫 토큰의 영문/숫자 및 선택적 하이픈 번호를 인식. 첫 공백 뒤 설명은 무시하며 영문 대소문자는 구분하지 않음."])
    info.append(["원본 보호", "원본 엑셀과 동영상은 수정/삭제하지 않았습니다."])

    for ws in wb.worksheets:
        style_sheet(ws)
    # 긴 파일 경로 열은 보기 편하게 넓히되 적절히 제한
    summary.column_dimensions["F"].width = 55
    missing.column_dimensions["F"].width = 55
    multiple.column_dimensions["F"].width = 55
    for ws in (summary, missing, multiple):
        for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
            row[0].alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(output_path)
    wb.close()
    return len(missing_rows), len(multiple_rows), len(by_id), len(no_id)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("관로번호 검사 프로그램")
        self.geometry("950x650")
        self.minsize(820, 560)
        self.excel_path = tk.StringVar()
        self.video_folder = tk.StringVar()
        self.output_path = tk.StringVar()
        self.recursive = tk.BooleanVar(value=True)
        self.last_report = None
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}
        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="관로번호 검사 프로그램", font=("Malgun Gothic", 16, "bold")).pack(anchor="w", pady=(0, 8))
        ttk.Label(main, text="엑셀의 '관로번호'와 동영상 파일명 시작 부분을 비교합니다. 원본 파일은 수정하지 않습니다.").pack(anchor="w", pady=(0, 12))

        file_frame = ttk.LabelFrame(main, text="검사 대상", padding=10)
        file_frame.pack(fill="x", **pad)
        self._path_row(file_frame, "엑셀 파일", self.excel_path, self.choose_excel, 0)
        self._path_row(file_frame, "동영상 폴더", self.video_folder, self.choose_folder, 1)
        self._path_row(file_frame, "결과 저장 위치", self.output_path, self.choose_output, 2)
        ttk.Checkbutton(file_frame, text="하위 폴더까지 모두 검색", variable=self.recursive).grid(row=3, column=1, sticky="w", pady=5)
        file_frame.columnconfigure(1, weight=1)

        action = ttk.Frame(main)
        action.pack(fill="x", **pad)
        self.run_button = ttk.Button(action, text="검사 실행", command=self.run_scan)
        self.run_button.pack(side="left")
        self.save_button = ttk.Button(action, text="결과 엑셀 저장", command=self.save_report, state="disabled")
        self.save_button.pack(side="left", padx=(8, 0))
        self.extra_button = ttk.Button(action, text="엑셀에 없는 영상 보기", command=self.show_extra_videos, state="disabled")
        self.extra_button.pack(side="left", padx=(8, 0))
        ttk.Button(action, text="초기화", command=self.clear_results).pack(side="left", padx=(8, 0))

        self.status = tk.StringVar(value="엑셀 파일과 동영상 폴더를 선택한 뒤 '검사 실행'을 누르세요.")
        ttk.Label(main, textvariable=self.status, foreground="#1f4e79", wraplength=900).pack(anchor="w", **pad)

        result_frame = ttk.LabelFrame(main, text="검사 결과 미리보기", padding=8)
        result_frame.pack(fill="both", expand=True, **pad)
        columns = ("id", "count", "status", "files")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        self.tree.heading("id", text="관로번호")
        self.tree.heading("count", text="동영상 개수")
        self.tree.heading("status", text="검사 결과")
        self.tree.heading("files", text="동영상 파일명")
        self.tree.column("id", width=150, anchor="center")
        self.tree.column("count", width=90, anchor="center")
        self.tree.column("status", width=120, anchor="center")
        self.tree.column("files", width=500, anchor="w")
        yscroll = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)
        self.tree.tag_configure("missing", foreground="#c00000")
        self.tree.tag_configure("multiple", foreground="#bf7000")
        self.tree.tag_configure("normal", foreground="#008000")

    def _path_row(self, parent, label, variable, command, row):
        ttk.Label(parent, text=label, width=14).grid(row=row, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(parent, text="찾아보기...", command=command).grid(row=row, column=2, padx=5, pady=5)

    def choose_excel(self):
        path = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
        if path:
            self.excel_path.set(path)
            if not self.output_path.get():
                self.output_path.set(str(Path(path).with_name(Path(path).stem + "_관로번호검사결과.xlsx")))

    def choose_folder(self):
        path = filedialog.askdirectory(title="동영상 폴더 선택")
        if path:
            self.video_folder.set(path)

    def choose_output(self):
        path = filedialog.asksaveasfilename(title="검사 결과 저장 위치", defaultextension=".xlsx",
                                            initialfile="관로번호검사결과.xlsx",
                                            filetypes=[("Excel 파일", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def run_scan(self):
        excel = self.excel_path.get().strip()
        folder = self.video_folder.get().strip()
        output = self.output_path.get().strip()
        if not excel or not os.path.isfile(excel):
            messagebox.showwarning("확인", "먼저 올바른 엑셀 파일을 선택해 주세요.")
            return
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("확인", "먼저 동영상 폴더를 선택해 주세요.")
            return
        if not output:
            messagebox.showwarning("확인", "결과 저장 위치를 지정해 주세요.")
            return
        if os.path.abspath(output).startswith(os.path.abspath(folder) + os.sep):
            if messagebox.askyesno("확인", "결과 파일을 동영상 폴더 안에 저장하려고 합니다. 결과 파일은 동영상이 아니므로 검사에는 영향이 없지만, 다른 위치에 저장하는 것을 권장합니다. 계속할까요?") is False:
                return

        self.run_button.configure(state="disabled")
        self.save_button.configure(state="disabled")
        self.status.set("검사 준비 중: 엑셀에서 관로번호 열을 찾고 있습니다...")
        self.update_idletasks()
        threading.Thread(target=self._scan_worker, args=(excel, folder, output, self.recursive.get()), daemon=True).start()

    def _scan_worker(self, excel, folder, output, recursive):
        try:
            self.after(0, lambda: self.status.set("1/2 엑셀의 실제 데이터 범위를 확인하고 관로번호를 읽는 중입니다..."))
            records, sheet_name, header_row, _ = find_header_and_rows(excel)
            self.after(0, lambda: self.status.set("2/2 영상 파일명 2,000개를 빠르게 확인하는 중입니다..."))
            by_id, no_id = scan_videos(folder, recursive)
            # GUI 조작은 메인 스레드에서 수행
            self.after(0, lambda: self._finish_scan(excel, folder, output, records, by_id, no_id, sheet_name, header_row))
        except Exception as exc:
            self.after(0, lambda msg=str(exc): self._scan_error(msg))

    def _finish_scan(self, excel, folder, output, records, by_id, no_id, sheet_name, header_row):
        self.last_report = {
            "excel": excel, "folder": folder, "output": output, "records": records,
            "by_id": by_id, "no_id": no_id, "sheet_name": sheet_name, "header_row": header_row
        }
        for item in self.tree.get_children():
            self.tree.delete(item)
        missing_count = 0
        multiple_count = 0
        preview_limit = 400
        for idx, rec in enumerate(records):
            paths = by_id.get(rec["id"], [])
            count = len(paths)
            if count == 0:
                status, tag = "영상 없음", "missing"
                missing_count += 1
            elif count > 1:
                status, tag = "영상 여러 개", "multiple"
                multiple_count += 1
            else:
                status, tag = "정상", "normal"
            # 화면 미리보기는 최대 400행만 표시해 창이 멈춘 것처럼 보이는 현상을 줄입니다.
            # 전체 데이터는 '결과 엑셀 저장'에 모두 포함됩니다.
            if idx < preview_limit:
                shown_files = " | ".join(Path(p).name for p in paths)
                self.tree.insert("", "end", values=(rec["id"], count, status, shown_files), tags=(tag,))
        video_count = sum(len(v) for v in by_id.values()) + len(no_id)
        excel_ids = {rec["id"] for rec in records}
        extra_ids = [code for code in by_id if code not in excel_ids]
        extra_video_count = sum(len(by_id[code]) for code in extra_ids)
        self.status.set(
            f"검사 완료: 엑셀 관로번호 {len(records)}행 / 검색한 동영상 {video_count}개 / "
            f"영상 없음 {missing_count}건 / 영상 여러 개 {multiple_count}건 / "
            f"엑셀에 없는 영상 {extra_video_count}개({len(extra_ids)}개 관로번호) / "
            f"번호 추출 불가 {len(no_id)}개. 전체 결과는 결과 엑셀에 저장됩니다."
        )
        self.run_button.configure(state="normal")
        self.save_button.configure(state="normal")
        self.extra_button.configure(state="normal")
        # 기본 경로가 비어 있지 않으면 바로 저장하지 않고 사용자가 저장 버튼을 누르게 함

    def show_extra_videos(self):
        """Show videos whose extracted IDs do not exist in the Excel ID column."""
        if not self.last_report:
            messagebox.showinfo("안내", "먼저 검사를 실행해 주세요.")
            return
        info = self.last_report
        excel_ids = {rec["id"] for rec in info["records"]}
        extras = []
        for code in sorted(info["by_id"].keys()):
            if code not in excel_ids:
                for path in info["by_id"][code]:
                    extras.append((code, path))

        for item in self.tree.get_children():
            self.tree.delete(item)
        if not extras:
            self.tree.insert("", "end", values=("없음", 0, "엑셀에 없는 영상 없음", "모든 번호가 엑셀 목록에 있습니다."))
            self.status.set("역방향 검사 완료: 엑셀 관로번호에 없는 동영상이 없습니다.")
            return

        for code, path in extras[:1000]:
            self.tree.insert("", "end", values=(code, 1, "엑셀에 없음", Path(path).name), tags=("missing",))
        self.status.set(
            f"역방향 검사 결과: 엑셀에 없는 동영상 {len(extras)}개. "
            f"화면에는 최대 1,000개까지 표시하며, 전체 목록은 결과 엑셀의 '엑셀에 없는 영상' 시트에서 확인하세요."
        )

    def save_report(self):
        if not self.last_report:
            messagebox.showinfo("안내", "먼저 검사를 실행해 주세요.")
            return
        info = self.last_report
        output = self.output_path.get().strip()
        if not output:
            output = filedialog.asksaveasfilename(title="검사 결과 저장", defaultextension=".xlsx",
                                                  initialfile="관로번호검사결과.xlsx",
                                                  filetypes=[("Excel 파일", "*.xlsx")])
            if not output:
                return
            self.output_path.set(output)
        if os.path.abspath(output) == os.path.abspath(info["excel"]):
            messagebox.showerror("저장 불가", "결과 저장 위치가 원본 엑셀과 같습니다. 원본 보호를 위해 다른 파일명을 선택해 주세요.")
            return
        try:
            missing, multiple, _, no_id = write_report(
                output, info["excel"], info["folder"], info["records"], info["by_id"], info["no_id"],
                info["sheet_name"], info["header_row"]
            )
            messagebox.showinfo("저장 완료", f"검사 결과를 저장했습니다.\n\n{output}\n\n영상 없음: {missing}건\n영상 여러 개: {multiple}건\n번호 추출 불가 영상: {no_id}개")
            self.status.set(f"결과 저장 완료: {output}")
        except PermissionError:
            messagebox.showerror("저장 실패", "결과 파일을 저장할 수 없습니다. 파일이 열려 있거나 폴더 권한이 없을 수 있습니다.")
        except Exception as exc:
            messagebox.showerror("저장 실패", f"결과 저장 중 오류가 발생했습니다.\n{exc}")

    def _scan_error(self, message):
        self.run_button.configure(state="normal")
        self.status.set("검사에 실패했습니다.")
        messagebox.showerror("검사 오류", message)

    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.last_report = None
        self.save_button.configure(state="disabled")
        self.extra_button.configure(state="disabled")
        self.status.set("초기화했습니다. 엑셀 파일과 동영상 폴더를 선택한 뒤 검사하세요.")


if __name__ == "__main__":
    if load_workbook is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("필수 패키지", "openpyxl이 설치되어 있지 않습니다.\n명령 프롬프트에서 다음 명령을 실행하세요:\n\npython -m pip install openpyxl")
    else:
        App().mainloop()
