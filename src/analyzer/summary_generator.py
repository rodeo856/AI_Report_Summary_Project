from collections import defaultdict
import json
from pathlib import path

class SummaryGenerator:
    def __init__(self, log_func=None, options=None):
        self.log = log_func or (lambda msg: None)
        self.options = options

    def generate(self, workbook, results):
        self.log("[SummaryGenerator] Status Sheet 생성 진입")

        self.delete_existing_summary(workbook)

        ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        ws.Name = "Status"
        ws.Activate()

        row = 1
        row = self.write_overall_summary(ws, row, results)
        row += 2

        row = self.write_daily_progress(ws, row, results)
        row += 2

        row = self.write_missing_items(ws, row, results)
        row += 2

        row = self.write_pass_fail_table(ws, row, results)

        self.merge_same_values_in_pass_fail(ws)
        self.apply_format(ws)
        self.format_pass_fail_text(ws)
        self.move_status_to_end(workbook)

        self.log("[SummaryGenerator] Status Sheet 생성 완료")

    def delete_existing_summary(self, workbook):
        app = workbook.Application

        for i in range(workbook.Worksheets.Count, 0, -1):
            ws = workbook.Worksheets(i)

            if ws.Name == "Status":
                app.DisplayAlerts = False
                ws.Delete()
                app.DisplayAlerts = True
                self.log("[SummaryGenerator] 기존 Status Sheet 삭제")
                break

    def write_title(self, ws, row, title):
        ws.Cells(row, 1).Value = title
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 1).Font.Size = 14

        # 제목 아래 공백 행
        ws.Rows(row + 1).RowHeight = 10
        return row + 2

    def write_overall_summary(self, ws, row, results):
        row = self.write_title(ws, row, "1. Summary")

        total_minor = len(results)
        waveform_target_minor = sum(1 for r in results if r.get("waveform_required_count", 0) > 0)
        ongoing_minor = sum(
            1 for r in results
            if 0 < r.get("waveform_done_count", 0) < r.get("waveform_required_count", 0)
        )
        complete_minor = sum(
            1 for r in results
            if r.get("waveform_required_count", 0) > 0
            and r.get("waveform_done_count", 0) >= r.get("waveform_required_count", 0)
        )

        total_slots = sum(r.get("waveform_total_slot_count", 0) for r in results)
        required_slots = sum(r.get("waveform_required_count", 0) for r in results)
        done_slots = sum(r.get("waveform_done_count", 0) for r in results)
        missing_slots = max(0, required_slots - done_slots)

        progress = done_slots / required_slots if required_slots else 0

        time_items = []
        for r in results:
            for t in r.get("waveform_time_list", []):
                time_items.append((t, r.get("sheet", "")))

        oldest_time, oldest_sheet = min(time_items) if time_items else ("", "")
        latest_time, latest_sheet = max(time_items) if time_items else ("", "")

        total_duration = self.format_duration(oldest_time, latest_time)

        ws.Cells(row, 1).Value = "소분류 항목"
        row += 1
        data = [
            ["전체 소분류 항목 수", total_minor],
            ["파형 대상 소분류 항목 수", waveform_target_minor],
            ["진행 중 소분류 항목 수", ongoing_minor],
            ["완료 소분류 항목 수", complete_minor],
        ]

        for label, value in data:
            ws.Cells(row, 1).Value = label
            ws.Cells(row, 2).Value = value
            row += 1

        ws.Rows(row).RowHeight = 10
        row += 1

        ws.Cells(row, 1).Value = "파형"
        row += 1
        data = [
            ["필요 파형 슬롯 수", required_slots],
            ["완료 파형 슬롯 수", done_slots],
            ["미완료 파형 슬롯 수", missing_slots],
            ["전체 파형 슬롯 수", total_slots],
        ]

        for label, value in data:
            ws.Cells(row, 1).Value = label
            ws.Cells(row, 2).Value = value
            row += 1

        ws.Rows(row).RowHeight = 10
        row += 1

        ws.Cells(row, 1).Value = "진행도"
        row += 1

        ws.Cells(row, 1).Value = "전체 진행률"
        ws.Cells(row, 2).Value = progress
        ws.Cells(row, 2).NumberFormat = "0.0%"
        row += 1

        ws.Cells(row, 1).Value = "시트 별 진행률"
        row += 1

        sheet_groups = defaultdict(list)
        for r in results:
            sheet_groups[r.get("sheet", "")].append(r)

        for sheet in sorted(sheet_groups.keys()):
            items = sheet_groups[sheet]
            s_required = sum(x.get("waveform_required_count", 0) for x in items)
            s_done = sum(x.get("waveform_done_count", 0) for x in items)
            s_rate = s_done / s_required if s_required else 0

            ws.Cells(row, 1).Value = sheet
            ws.Cells(row, 2).Value = s_rate
            ws.Cells(row, 2).NumberFormat = "0.0%"
            row += 1

        ws.Cells(row, 1).Value = "최초 측정일"
        ws.Cells(row, 2).Value = f"{oldest_time} / {oldest_sheet}" if oldest_time else ""
        row += 1

        ws.Cells(row, 1).Value = "최신 측정일"
        ws.Cells(row, 2).Value = f"{latest_time} / {latest_sheet}" if latest_time else ""
        row += 1

        ws.Cells(row, 1).Value = "전체 측정 시간"
        ws.Cells(row, 2).Value = total_duration
        row += 1

        return row

    def write_daily_progress_all(self, ws, row, results):
        row = self.write_title(ws, row, "2. Daily Progress - All Sheets")

        headers = ["날짜", "일일 진행 항목 수", "누적 진행 항목 수", "전체 항목 수", "누적 진행률"]
        row = self.write_header(ws, row, headers)

        daily = defaultdict(int)

        for r in results:
            if r.get("status") != "진행":
                continue

            t = r.get("latest_waveform_time", "")
            if not t:
                continue

            daily[t[:10]] += 1

        total = len(results)
        cumulative = 0

        for date in sorted(daily.keys()):
            count = daily[date]
            cumulative += count

            ws.Cells(row, 1).Value = date
            ws.Cells(row, 2).Value = count
            ws.Cells(row, 3).Value = cumulative
            ws.Cells(row, 4).Value = total
            ws.Cells(row, 5).Value = cumulative / total if total else 0
            ws.Cells(row, 5).NumberFormat = "0.0%"
            row += 1

        return row

    def write_daily_progress_by_sheet(self, ws, row, results):
        row = self.write_title(ws, row, "3. Daily Progress - By Sheet")

        headers = ["Sheet", "날짜", "일일 진행 항목 수", "누적 진행 항목 수", "Sheet 전체 항목 수", "누적 진행률"]
        row = self.write_header(ws, row, headers)

        sheet_groups = defaultdict(list)
        for r in results:
            sheet_groups[r.get("sheet", "")].append(r)

        for sheet in sorted(sheet_groups.keys()):
            items = sheet_groups[sheet]
            daily = defaultdict(int)

            for r in items:
                if r.get("status") != "진행":
                    continue

                t = r.get("latest_waveform_time", "")
                if not t:
                    continue

                daily[t[:10]] += 1

            total = len(items)
            cumulative = 0

            for date in sorted(daily.keys()):
                count = daily[date]
                cumulative += count

                ws.Cells(row, 1).Value = sheet
                ws.Cells(row, 2).Value = date
                ws.Cells(row, 3).Value = count
                ws.Cells(row, 4).Value = cumulative
                ws.Cells(row, 5).Value = total
                ws.Cells(row, 6).Value = cumulative / total if total else 0
                ws.Cells(row, 6).NumberFormat = "0.0%"
                row += 1

        return row

    def write_missing_items(self, ws, row, results):
        row = self.write_title(ws, row, "3. Missing / Not Completed Items")

        headers = [
            "No.",
            "Sheet",
            "대분류",
            "중분류",
            "소분류",
            "State",
        ]
        row = self.write_header(ws, row, headers)

        count = 1

        for r in results:
            done_text = "Done" if r.get("status") == "진행" else "On going"

            if done_text == "Done":
                continue

            values = [
                count,
                r.get("sheet", ""),
                r.get("major", ""),
                r.get("middle", ""),
                r.get("minor", ""),
                done_text,
            ]

            for col, value in enumerate(values, start=1):
                ws.Cells(row, col).Value = value

            count += 1
            row += 1

        return row

    def write_pass_fail_table(self, ws, row, results):
        row = self.write_title(ws, row, "4. Pass/Fail Result")

        headers = [
            "Sheet",
            "대분류",
            "중분류",
            "소분류",
            "Pass/Fail",
        ]
        row = self.write_header(ws, row, headers)

        for r in results:
            values = [
                r.get("sheet", ""),
                r.get("major", ""),
                r.get("middle", ""),
                r.get("minor", ""),
                r.get("pass_fail", ""),
            ]

            for col, value in enumerate(values, start=1):
                ws.Cells(row, col).Value = value

            row += 1

        return row

    def merge_same_values_in_pass_fail(self, ws):
        xlCenter = -4108
        xlLeft = -4131

        used = ws.UsedRange
        last_row = used.Row + used.Rows.Count - 1

        start_row = None
        for row in range(1, last_row + 1):
            value = ws.Cells(row, 1).Value
            value = "" if value is None else str(value).strip()

            if value.startswith("4. Pass/Fail"):
                start_row = row
                break

        if start_row is None:
            return

        header_row = start_row + 1
        data_start = header_row + 1

        data_end = data_start - 1
        for row in range(data_start, last_row + 1):
            sheet = ws.Cells(row, 1).Value
            major = ws.Cells(row, 2).Value
            middle = ws.Cells(row, 3).Value
            minor = ws.Cells(row, 4).Value

            if any(v not in [None, ""] for v in [sheet, major, middle, minor]):
                data_end = row

        if data_end < data_start:
            return

        # 병합 전 값 캐시
        cache = {}
        for row in range(data_start, data_end + 1):
            cache[row] = {
                1: self.normalize_text(ws.Cells(row, 1).Value),
                2: self.normalize_text(ws.Cells(row, 2).Value),
                3: self.normalize_text(ws.Cells(row, 3).Value),
            }

        app = ws.Application
        old_alerts = app.DisplayAlerts
        app.DisplayAlerts = False

        try:
            # 기존 병합 해제
            ws.Range(ws.Cells(data_start, 1), ws.Cells(data_end, 3)).UnMerge()

            # 아래 계층부터 병합
            self.merge_runs_from_cache(ws, cache, data_start, data_end, 3, [1, 2, 3])
            self.merge_runs_from_cache(ws, cache, data_start, data_end, 2, [1, 2])
            self.merge_runs_from_cache(ws, cache, data_start, data_end, 1, [1])

            # 정렬
            rng = ws.Range(ws.Cells(data_start, 1), ws.Cells(data_end, 3))
            rng.VerticalAlignment = xlCenter
            rng.HorizontalAlignment = xlLeft

        finally:
            app.DisplayAlerts = old_alerts

    def merge_runs_from_cache(self, ws, cache, start_row, end_row, target_col, key_cols):
        xlCenter = -4108
        xlLeft = -4131

        run_start = start_row
        prev_key = tuple(cache[start_row][c] for c in key_cols)

        for row in range(start_row + 1, end_row + 2):
            if row <= end_row:
                cur_key = tuple(cache[row][c] for c in key_cols)
            else:
                cur_key = None

            if cur_key != prev_key:
                if row - run_start > 1:
                    rng = ws.Range(
                        ws.Cells(run_start, target_col),
                        ws.Cells(row - 1, target_col)
                    )
                    rng.Merge()
                    rng.VerticalAlignment = xlCenter
                    rng.HorizontalAlignment = xlLeft

                run_start = row
                prev_key = cur_key

    def normalize_text(self, value):
        if value is None:
            return ""

        text = str(value).strip()
        text = text.replace("\n", " ")
        text = text.replace("\r", " ")
        text = text.replace("\t", " ")
        text = " ".join(text.split())
        return text

    def format_pass_fail_text(self, ws):
        used = ws.UsedRange
        last_row = used.Row + used.Rows.Count - 1
        last_col = used.Column + used.Columns.Count - 1

        pass_fail_col = None
        start_row = None

        for row in range(1, last_row + 1):
            for col in range(1, last_col + 1):
                value = ws.Cells(row, col).Value
                value = "" if value is None else str(value).strip()

                if value == "Pass/Fail":
                    pass_fail_col = col
                    start_row = row + 1
                    break

            if pass_fail_col:
                break

        if not pass_fail_col:
            return

        for row in range(start_row, last_row + 1):
            cell = ws.Cells(row, pass_fail_col)
            text = cell.Value

            if text is None:
                continue

            text = str(text).strip()

            if not text:
                continue

            # 기본값
            cell.Font.Color = self.rgb(0, 0, 0)
            cell.Font.Bold = False

            parts = text.split("/")

            pos = 1

            for part in parts:

                if part == "FAIL":

                    try:
                        chars = cell.GetCharacters(pos, len(part))
                        chars.Font.Color = self.rgb(255, 0, 0)
                        chars.Font.Bold = True
                    except Exception as e:
                        print(e)

                pos += len(part) + 1

    def apply_format(self, ws):
        xlCenter = -4108
        xlLeft = -4131
        xlContinuous = 1
        xlNone = -4142

        used = ws.UsedRange
        last_row = used.Row + used.Rows.Count - 1
        last_col = used.Column + used.Columns.Count - 1

        used.Font.Name = "맑은 고딕"
        used.Font.Size = 10
        used.Font.Bold = False
        used.Font.Color = 0x000000
        used.Interior.Pattern = xlNone
        used.Borders.LineStyle = xlNone

        main_title_rows = []
        sub_title_rows = []
        summary_sub_title_rows = []
        info_rows = []
        header_rows = []

        for row in range(1, last_row + 1):
            value = ws.Cells(row, 1).Value
            if value is None:
                continue

            value = str(value).strip()

            if value.startswith((
                "1. Summary",
                "2. Daily",
                "3. Missing",
                "4. Pass/Fail",
            )):
                main_title_rows.append(row)

            elif value in ["소분류 항목", "파형", "진행도"]:
                summary_sub_title_rows.append(row)

            elif value.startswith((
                "2-1)",
                "2-2)",
            )):
                sub_title_rows.append(row)

            elif value in ["전체 측정 시간"] and row > 12:
                info_rows.append(row)

            if value in ["Sheet", "날짜", "구분", "No."]:
                header_rows.append(row)

        for row in main_title_rows:
            rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, last_col))
            rng.Font.Bold = True
            rng.Font.Size = 12
            rng.Font.Color = 0xFFFFFF
            rng.Interior.Color = self.rgb(84, 130, 53)
            rng.HorizontalAlignment = xlLeft
            rng.Borders.LineStyle = xlNone

        for row in sub_title_rows:
            rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, last_col))
            rng.Font.Bold = True
            rng.Font.Size = 11
            rng.Font.Color = 0x000000
            rng.Interior.Color = self.rgb(226, 240, 217)
            rng.HorizontalAlignment = xlLeft
            rng.Borders.LineStyle = xlNone

        for row in summary_sub_title_rows:
            rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, last_col))
            rng.Font.Bold = True
            rng.Font.Size = 11
            rng.Font.Color = 0x000000
            rng.Interior.Pattern = xlNone
            rng.HorizontalAlignment = xlLeft
            rng.Borders.LineStyle = xlNone

        for row in info_rows:
            rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, last_col))
            rng.Font.Bold = False
            rng.Font.Color = 0x000000
            rng.Interior.Pattern = xlNone
            rng.Borders.LineStyle = xlNone

        for row in header_rows:
            header_last_col = 0
            for col in range(last_col, 0, -1):
                value = ws.Cells(row, col).Value
                if value not in [None, ""]:
                    header_last_col = col
                    break

            if header_last_col == 0:
                continue

            rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, header_last_col))
            rng.Font.Bold = True
            rng.Font.Size = 10
            rng.Font.Color = 0x000000
            rng.Interior.Color = self.rgb(221, 235, 247)
            rng.HorizontalAlignment = xlCenter

        no_border_rows = set(main_title_rows + sub_title_rows + summary_sub_title_rows + info_rows)

        current_section = ""

        for row in range(1, last_row + 1):
            value_a = ws.Cells(row, 1).Value
            value_a = "" if value_a is None else str(value_a).strip()

            if value_a.startswith("4. Pass/Fail"):
                current_section = "pass_fail"
            elif value_a.startswith(("1. Summary", "2. Daily", "3. Missing")):
                current_section = ""

            if row in no_border_rows:
                continue

            row_last_col = 0
            for col in range(last_col, 0, -1):
                value = ws.Cells(row, col).Value
                if value not in [None, ""]:
                    row_last_col = col
                    break

            if row_last_col == 0:
                continue

            if current_section == "pass_fail":
                row_last_col = max(row_last_col, 5)

            for col in range(1, row_last_col + 1):
                ws.Cells(row, col).Borders.LineStyle = xlContinuous

        # Summary 영역의 "전체 측정 시간" 행만 테두리 강제 적용
        for row in range(1, last_row + 1):
            value = ws.Cells(row, 1).Value
            value = "" if value is None else str(value).strip()

            if value == "전체 측정 시간" and row < 22:
                rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, 2))
                rng.Borders.LineStyle = xlContinuous
                break

        ws.Columns.AutoFit()

        for col in range(1, last_col + 1):
            if ws.Columns(col).ColumnWidth > 45:
                ws.Columns(col).ColumnWidth = 45

        ws.Activate()
        ws.Application.ActiveWindow.FreezePanes = False
        ws.Application.ActiveWindow.SplitRow = 0
        ws.Application.ActiveWindow.SplitColumn = 0

    def write_header(self, ws, row, headers):
        for col, h in enumerate(headers, start=1):
            ws.Cells(row, col).Value = h
            ws.Cells(row, col).Font.Bold = True
        return row + 1

    def write_category_progress(self, ws, row, results):
        row = self.write_title(ws, row, "5. Category Progress")

        headers = ["구분", "항목", "전체 수", "진행 수", "진행률"]
        row = self.write_header(ws, row, headers)

        groups = {
            "Operation": "operation",
            "Input": "input",
            "Ratio": "ratio",
            "대분류": "major",
            "중분류": "middle",
        }

        for group_name, key in groups.items():
            grouped = defaultdict(list)

            for r in results:
                grouped[r.get(key, "")].append(r)

            for item_name, items in grouped.items():
                if not item_name:
                    continue

                total = len(items)
                done = sum(1 for x in items if x.get("status") == "진행")
                rate = done / total if total else 0

                ws.Cells(row, 1).Value = group_name
                ws.Cells(row, 2).Value = item_name
                ws.Cells(row, 3).Value = total
                ws.Cells(row, 4).Value = done
                ws.Cells(row, 5).Value = rate
                ws.Cells(row, 5).NumberFormat = "0.0%"
                row += 1

        return row

    def write_detail_table(self, ws, row, results):
        row = self.write_title(ws, row, "6. Detail Table")

        headers = [
            "No.",
            "Sheet",
            "Operation",
            "Input",
            "Ratio",
            "대분류",
            "중분류",
            "소분류",
            "Start Row",
            "End Row",
            "Waveform O Count",
            "Waveform X Count",
            "Waveform File Count",
            "Latest Waveform Time",
            "Status",
            "Waveform Files",
        ]

        row = self.write_header(ws, row, headers)

        for idx, r in enumerate(results, start=1):
            values = [
                idx,
                r.get("sheet", ""),
                r.get("operation", ""),
                r.get("input", ""),
                r.get("ratio", ""),
                r.get("major", ""),
                r.get("middle", ""),
                r.get("minor", ""),
                r.get("start_row", ""),
                r.get("end_row", ""),
                r.get("waveform_o_count", 0),
                r.get("waveform_x_count", 0),
                r.get("waveform_file_count", 0),
                r.get("latest_waveform_time", ""),
                r.get("status", ""),
                r.get("waveform_files", ""),
            ]

            for col, value in enumerate(values, start=1):
                ws.Cells(row, col).Value = value

            row += 1

    def format_duration(self, start_time, end_time):
        if not start_time or not end_time:
            return ""

        from datetime import datetime

        try:
            if isinstance(start_time, str):
                start_time = datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S")
            if isinstance(end_time, str):
                end_time = datetime.strptime(end_time, "%Y-%m-%d %H:%M:%S")

            delta = end_time - start_time
        except Exception:
            return ""

        total_minutes = int(delta.total_seconds() // 60)

        days = total_minutes // 1440
        hours = (total_minutes % 1440) // 60
        minutes = total_minutes % 60

        total_hours = total_minutes // 60
        total_rem_minutes = total_minutes % 60

        return f"{days} days / {hours} hrs / {minutes} min ({total_hours} hrs {total_rem_minutes} min)"
    
    def write_daily_progress(self, ws, row, results):
        from collections import defaultdict

        row = self.write_title(ws, row, "2. Daily Progress")

        all_times = []
        for r in results:
            all_times.extend(r.get("waveform_time_list", []))

        row = self.write_title(ws, row, "2-1) Total")

        if all_times:
            ws.Cells(row, 1).Value = "전체 측정 시간"
            ws.Cells(row, 2).Value = self.format_duration(min(all_times), max(all_times))
            row += 1

        headers = ["날짜", "일일 진행 항목 수", "누적 진행 항목 수", "전체 항목 수", "누적 진행률"]
        row = self.write_header(ws, row, headers)

        daily = defaultdict(int)
        for r in results:
            if r.get("status") != "진행":
                continue
            t = r.get("latest_waveform_time", "")
            if t:
                daily[t[:10]] += 1

        total = len(results)
        cumulative = 0

        for date in sorted(daily.keys()):
            count = daily[date]
            cumulative += count
            ws.Cells(row, 1).Value = date
            ws.Cells(row, 2).Value = count
            ws.Cells(row, 3).Value = cumulative
            ws.Cells(row, 4).Value = total
            ws.Cells(row, 5).Value = cumulative / total if total else 0
            ws.Cells(row, 5).NumberFormat = "0.0%"
            row += 1

        row += 2
        row = self.write_title(ws, row, "2-2) Sheet 별")

        sheet_groups = defaultdict(list)
        for r in results:
            sheet_groups[r.get("sheet", "")].append(r)

        headers = ["Sheet", "측정 시간", "날짜", "일일 진행 항목 수", "누적 진행 항목 수", "Sheet 전체 항목 수", "누적 진행률"]
        row = self.write_header(ws, row, headers)

        for sheet in sorted(sheet_groups.keys()):
            items = sheet_groups[sheet]

            times = []
            for r in items:
                times.extend(r.get("waveform_time_list", []))

            sheet_duration = self.format_duration(min(times), max(times)) if times else ""

            daily = defaultdict(int)
            for r in items:
                if r.get("status") != "진행":
                    continue
                t = r.get("latest_waveform_time", "")
                if t:
                    daily[t[:10]] += 1

            total = len(items)
            cumulative = 0

            for date in sorted(daily.keys()):
                count = daily[date]
                cumulative += count

                ws.Cells(row, 1).Value = sheet
                ws.Cells(row, 2).Value = sheet_duration
                ws.Cells(row, 3).Value = date
                ws.Cells(row, 4).Value = count
                ws.Cells(row, 5).Value = cumulative
                ws.Cells(row, 6).Value = total
                ws.Cells(row, 7).Value = cumulative / total if total else 0
                ws.Cells(row, 7).NumberFormat = "0.0%"
                row += 1

        return row

    def rgb(self, r, g, b):
        return r + (g * 256) + (b * 65536)
    
    def move_status_to_end(self, workbook):
        try:
            ws = workbook.Worksheets("Status")

            # 이미 맨 끝이면 Move 하지 않음
            if ws.Index == workbook.Worksheets.Count:
                self.log("[SummaryGenerator] Status Sheet 이미 맨 끝")
                ws.Activate()
                return

            ws.Move(After=workbook.Worksheets(workbook.Worksheets.Count))
            workbook.Worksheets("Status").Activate()
            self.log("[SummaryGenerator] Status Sheet 맨 끝 이동 완료")

        except Exception as e:
            self.log(f"[SummaryGenerator] Status Sheet 맨 끝 이동 실패: {e}")

