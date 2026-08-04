import re
from datetime import datetime


WAVEFORM_FILE_RE = re.compile(r"M_(\d{8})_(\d{6})\.png", re.IGNORECASE)


class WaveformParser:
    def find_waveform_files(self, text: str):
        if not text:
            return []
        return [m.group(0) for m in WAVEFORM_FILE_RE.finditer(str(text))]

    def parse_waveform_time(self, filename: str):
        if not filename:
            return ""
        m = WAVEFORM_FILE_RE.search(str(filename))
        if not m:
            return ""
        try:
            dt = datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S")
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            return ""
