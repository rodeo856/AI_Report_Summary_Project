import re
from src.analyzer.waveform_parser import WaveformParser
from src.analyzer.image_time_resolver import ImageTimeResolver


TARGET_SHEET_RE = re.compile(r"^(RVS|FWD)_[VW][1-4]$")


class ReportParser:
    def __init__(self, log_func=None):
        self.log = log_func or (lambda msg: None)
        self.waveform_parser = WaveformParser()
        self.image_time_resolver = ImageTimeResolver(self.log)

    def parse_workbook(self, workbook):
        results = []
        self.log("[ReportParser] parse_workbook 진입")

        for ws in workbook.Worksheets:
            sheet_name = str(ws.Name)
            if not TARGET_SHEET_RE.match(sheet_name):
                self.log(f"[ReportParser] Skip: {sheet_name}")
                continue

            self.log(f"분석 중: {sheet_name}")
            sheet_results = self.parse_sheet(ws)
            results.extend(sheet_results)
            self.log(f"[ReportParser] 완료: {sheet_name} / {len(sheet_results)}개 항목")

        self.log(f"[ReportParser] 전체 파싱 완료 / 총 {len(results)}개 항목")
        return results

    def parse_sheet(self, ws):
        sheet_name = str(ws.Name)
        used = ws.UsedRange
        start_row = used.Row
        values = used.Value

        if values is None:
            return []
        if not isinstance(values, tuple):
            values = ((values,),)
        elif values and not isinstance(values[0], tuple):
            values = (values,)

        row_count = len(values)
        col_count = len(values[0]) if row_count else 0
        self.log(f"[ReportParser] {sheet_name} UsedRange: start_row={start_row}, rows={row_count}, cols={col_count}")

        categories = self.find_category_rows(values, start_row)
        self.log(f"[ReportParser] {sheet_name} 카테고리 수: {len(categories)}")

        op, rest = sheet_name.split("_")
        input_code = rest[0]
        ratio = rest[1]
        input_name = "VBUS" if input_code == "V" else "WPC"

        results = []
        current_major = ""
        current_middle = ""

        for idx, item in enumerate(categories):
            if item["level"] == "major":
                current_major = item["name"]
                current_middle = ""
                continue
            if item["level"] == "middle":
                current_middle = item["name"]
                continue
            if item["level"] != "minor":
                continue

            start = item["row"]
            next_rows = [
                x["row"] for x in categories[idx + 1:]
                if x["level"] in ["major", "middle", "minor"]
            ]
            end = min(next_rows) - 1 if next_rows else start_row + row_count - 1

            analyzed = self.analyze_minor_section(ws, values, start_row, start, end)

            results.append({
                "sheet": sheet_name,
                "operation": op,
                "input": input_name,
                "ratio": ratio,
                "major": current_major,
                "middle": current_middle,
                "minor": item["name"],
                "start_row": start,
                "end_row": end,
                "waveform_o_count": analyzed["waveform_o_count"],
                "waveform_x_count": analyzed["waveform_x_count"],
                "waveform_file_count": analyzed["waveform_file_count"],
                "latest_waveform_time": analyzed["latest_waveform_time"],
                "status": analyzed["status"],
                "waveform_files": analyzed["waveform_files"],
                "waveform_file_list": analyzed.get("waveform_file_list", []),
                "waveform_time_list": analyzed.get("waveform_time_list", []),
                "pass_fail": analyzed.get("pass_fail", ""),
                "pass_fail_list": analyzed.get("pass_fail_list", []),
                "waveform_total_slot_count": analyzed.get("waveform_total_slot_count", 0),
                "waveform_excluded_count": analyzed.get("waveform_excluded_count", 0),
                "waveform_required_count": analyzed.get("waveform_required_count", 0),
                "waveform_done_count": analyzed.get("waveform_done_count", 0),
                "waveform_missing_count": analyzed.get("waveform_missing_count", 0),
                "image_shape_count": analyzed.get("image_shape_count",0,),
                "image_link_count": analyzed.get("image_link_count",0,),
                "image_link_paths": analyzed.get("image_link_paths",[],),
                "image_root_match_count": analyzed.get("image_root_match_count",0,),
                "image_metadata_time_list": analyzed.get("image_metadata_time_list",[],),
                "image_metadata_records": analyzed.get("image_metadata_records",[],),
                "latest_image_metadata_time": analyzed.get("latest_image_metadata_time","",),
                "progress_state": analyzed.get("progress_state", ""),
                "waveform_progress_rate": analyzed.get("waveform_progress_rate", ""),
                })
        return results

    def find_category_rows(self, values, used_start_row):
        rows = []

        def get_row_text(row_values):
            cells = [self.cell(row_values, col) for col in range(1, 12)]
            text = " ".join([x for x in cells if x]).strip()
            return re.sub(r"\s+", " ", text)

        def next_non_empty_text(start_idx):
            for j in range(start_idx + 1, min(start_idx + 8, len(values))):
                text = get_row_text(values[j])
                if text:
                    return text
            return ""

        def normalize_category_name(level, raw_text):
            text = re.sub(r"\s+", " ", raw_text).strip()
            text = text.replace(".0 ", " ")

            if level == "major":
                m = re.match(r"^(\d+)(?:\.0)?\s+(.+)", text)
                if m:
                    return f"{m.group(1)}. {m.group(2)}"

            if level == "middle":
                m = re.match(r"^(\d+)(?:\.0)?\)?\s+(.+)", text)
                if m:
                    return f"{m.group(1)}) {m.group(2)}"

            if level == "minor":
                m = re.match(r"^([A-Z])\.?\s+(.+)", text)
                if m:
                    return f"{m.group(1)}. {m.group(2)}"

            return text

        for i, row_values in enumerate(values):
            excel_row = used_start_row + i
            row_text = get_row_text(row_values)

            if not row_text:
                continue

            next_text = next_non_empty_text(i)

            # 소분류: A VOUT POK
            if re.match(r"^[A-Z]\s+[A-Za-z가-힣].+", row_text):
                rows.append({
                    "row": excel_row,
                    "level": "minor",
                    "name": normalize_category_name("minor", row_text),
                })
                continue

            # 중분류:
            # 1) Qualification Process 또는 Excel이 1 Qualification Process로 읽은 경우
            # 단, 바로 아래쪽에 소분류 A/B/C가 이어져야 함
            if re.match(r"^\d+(\.0)?\)?\s+[A-Za-z가-힣].+", row_text):
                if re.match(r"^[A-Z]\s+[A-Za-z가-힣].+", next_text):
                    rows.append({
                        "row": excel_row,
                        "level": "middle",
                        "name": normalize_category_name("middle", row_text),
                    })
                    continue

            # 대분류:
            # 숫자 + 제목이고, 바로 아래가 중분류 패턴이어야 함
            if re.match(r"^\d+(\.0)?\s+[A-Za-z가-힣].+", row_text):
                if re.match(r"^\d+(\.0)?\)?\s+[A-Za-z가-힣].+", next_text):
                    rows.append({
                        "row": excel_row,
                        "level": "major",
                        "name": normalize_category_name("major", row_text),
                    })
                    continue

        return rows

    def analyze_minor_section(self, ws, values, used_start_row, start_row, end_row):
        waveform_o_count = 0
        waveform_x_count = 0
        waveform_files = []
        waveform_times = []

        start_idx = max(start_row - used_start_row, 0)
        end_idx = min(end_row - used_start_row, len(values) - 1)

        for i in range(start_idx, end_idx + 1):
            row_values = values[i]
            l_val = self.cell(row_values, 12)
            m_val = self.cell(row_values, 13).upper()

            if "WAVEFORM" in l_val.upper():
                if m_val == "O":
                    waveform_o_count += 1
                elif m_val == "X":
                    waveform_x_count += 1

            for cell_value in row_values:
                text = "" if cell_value is None else str(cell_value).strip()
                if not text:
                    continue
                for f in self.waveform_parser.find_waveform_files(text):
                    waveform_files.append(f)
                    dt = self.waveform_parser.parse_waveform_time(f)
                    if dt:
                        waveform_times.append(dt)

        image_time_info = self.image_time_resolver.inspect_section(
            ws=ws,
            start_row=start_row,
            end_row=end_row,
            waveform_files=waveform_files,
            )

        image_shape_count = image_time_info["image_shape_count"]

        image_metadata_times = image_time_info["metadata_time_list"]

        latest_image_metadata_time = (
            max(image_metadata_times)
            if image_metadata_times
            else ""
        )
        waveform_total_slot_count = waveform_o_count

        waveform_excluded_count = 0

        waveform_required_count = max(
            0,
            waveform_total_slot_count - waveform_excluded_count
        )

        png_count = len(set(waveform_files))

        waveform_done_count = min(
            waveform_required_count,
            max(png_count, image_shape_count)
        )

        waveform_missing_count = max(
            0,
            waveform_required_count - waveform_done_count
        )

        if waveform_required_count == 0:
            progress_state = "No Target"
            status = "파형 미대상"

        elif waveform_done_count >= waveform_required_count:
            progress_state = "Done"
            status = "진행"

        elif waveform_done_count > 0:
            progress_state = "On going"
            status = "진행중"

        else:
            progress_state = "Not started"
            status = "미진행"

        pass_fail_values = []

        for i in range(start_idx, end_idx + 1):
            row_values = values[i]
            row_text = " ".join(
                str(v).strip() for v in row_values if v is not None
            ).upper()

            if "TEST RESULT SUMMARY" in row_text:
                for j in range(i, min(i + 15, end_idx + 1)):
                    header_row = values[j]

                    pass_fail_col = None
                    for col_idx, cell_value in enumerate(header_row):
                        if cell_value and str(cell_value).strip().upper() in ["PASS/FAIL", "PASS FAIL"]:
                            pass_fail_col = col_idx
                            break

                    if pass_fail_col is not None:
                        for k in range(j + 1, end_idx + 1):
                            v = values[k][pass_fail_col] if pass_fail_col < len(values[k]) else None
                            text = "" if v is None else str(v).strip().upper()

                            if text == "PASS":
                                pass_fail_values.append("PASS")
                            elif text == "FAIL":
                                pass_fail_values.append("FAIL")

                        break

        pass_fail_list = self.extract_pass_fail_list(values, start_idx, end_idx) 
        pass_fail_list = pass_fail_list[:waveform_done_count]

        latest_time = max(waveform_times) if waveform_times else ""
        return {
            "waveform_o_count": waveform_o_count,
            "waveform_x_count": waveform_x_count,
            "waveform_file_count": len(waveform_files),
            "latest_waveform_time": latest_time,
            "status": status,
            "waveform_files": ", ".join(waveform_files),
            "waveform_file_list": waveform_files,
            "waveform_time_list": waveform_times,
            "pass_fail": "FAIL" if "FAIL" in pass_fail_values else ("PASS" if "PASS" in pass_fail_values else ""),
            "pass_fail_list": pass_fail_list,
            "pass_fail": "/".join(pass_fail_list),
            "waveform_total_slot_count": waveform_total_slot_count,
            "waveform_excluded_count": waveform_excluded_count,
            "waveform_required_count": waveform_required_count,
            "waveform_done_count": waveform_done_count,
            "waveform_missing_count": waveform_missing_count,
            "image_shape_count": image_shape_count,
            "image_link_count": image_time_info["linked_picture_count"],
            "image_link_paths": image_time_info["linked_paths"],

            "image_root_match_count": (
                image_time_info["image_root_match_count"]
            ),

            "image_metadata_time_list": (
                image_time_info["metadata_time_list"]
            ),

            "image_metadata_records": (
                image_time_info["metadata_records"]
            ),

            "latest_image_metadata_time": (
                latest_image_metadata_time
            ),
            "progress_state": progress_state,
            "waveform_progress_rate":
                waveform_done_count / waveform_required_count
                if waveform_required_count else 0,
        }

    def count_image_shapes_in_section(self, ws, start_row, end_row):
        count = 0

        try:
            shapes = ws.Shapes
        except Exception:
            return 0

        for shape in shapes:
            try:
                top_row = shape.TopLeftCell.Row
                bottom_row = shape.BottomRightCell.Row

                if bottom_row < start_row or top_row > end_row:
                    continue

                shape_type = int(shape.Type)

                # 13: msoPicture, 11: msoLinkedPicture
                if shape_type in [11, 13]:
                    count += 1

            except Exception:
                continue

        return count

    @staticmethod
    def cell(row_values, excel_col):
        idx = excel_col - 1
        if idx < 0 or idx >= len(row_values):
            return ""
        v = row_values[idx]
        return "" if v is None else str(v).strip()

    def extract_pass_fail_list(self, values, start_idx, end_idx):
        pass_fail_list = []

        for i in range(start_idx, end_idx + 1):
            row_values = values[i]

            row_text = " ".join(
                "" if v is None else str(v).strip()
                for v in row_values
            ).upper()

            if "TEST RESULT SUMMARY" not in row_text:
                continue

            # TEST RESULT SUMMARY 아래쪽에서 PASS/FAIL 헤더 찾기
            for header_i in range(i + 1, min(i + 20, end_idx + 1)):
                header_values = values[header_i]

                pass_fail_col = None

                for col_idx, cell_value in enumerate(header_values):
                    text = "" if cell_value is None else str(cell_value).strip().upper()

                    if text in ["PASS/FAIL", "PASS FAIL", "PASS_FAIL"]:
                        pass_fail_col = col_idx
                        break

                if pass_fail_col is None:
                    continue

                # 헤더 아래 데이터 영역에서 PASS/FAIL 값 수집
                for data_i in range(header_i + 1, end_idx + 1):
                    data_values = values[data_i]

                    if pass_fail_col >= len(data_values):
                        continue

                    text = "" if data_values[pass_fail_col] is None else str(data_values[pass_fail_col]).strip().upper()

                    if text in ["PASS", "FAIL"]:
                        pass_fail_list.append(text)

                    # 다음 섹션으로 넘어간 것 같으면 중단
                    row_text2 = " ".join(
                        "" if v is None else str(v).strip()
                        for v in data_values
                    ).upper()

                    if "CONTENTS" in row_text2:
                        break

                return pass_fail_list

        return pass_fail_list