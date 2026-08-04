from dataclasses import dataclass


@dataclass
class AnalyzeOptions:
    recreate_summary: bool = True
    waveform_check: bool = True
    daily_progress: bool = True
    data_table_check: bool = False
    image_exist_check: bool = False
