import os
import re
import struct
from datetime import datetime, timedelta
from pathlib import Path

from PIL import ExifTags, Image


class ImageTimeResolver:
    """Resolve waveform image source paths and read time metadata.

    This helper is intentionally read-only. It never modifies the workbook or PNG files.
    """

    DEFAULT_IMAGE_ROOT = Path(r"D:\tek_remote_image")
    PICTURE_TYPES = {11, 13}  # msoLinkedPicture, msoPicture
    LINKED_PICTURE_TYPE = 11

    TEXT_TIME_KEYS = {
        "creationtime",
        "datetimeoriginal",
        "datetimedigitized",
        "datetime",
        "timestamp",
        "capturetime",
        "acquisitiontime",
    }

    EXIF_TIME_PRIORITY = {
        "DateTimeOriginal": 10,
        "DateTimeDigitized": 20,
        "DateTime": 30,
    }

    def __init__(self, log_func=None, image_root=None):
        self.log = log_func or (lambda message: None)
        self.image_root = Path(image_root) if image_root else self.DEFAULT_IMAGE_ROOT
        self._png_index = None

    def inspect_section(self, ws, start_row, end_row, waveform_files):
        result = {
            "image_shape_count": 0,
            "linked_picture_count": 0,
            "linked_paths": [],
            "image_root_match_count": 0,
            "metadata_time_list": [],
            "metadata_records": [],
        }

        seen_paths = set()

        try:
            shapes = ws.Shapes
        except Exception:
            shapes = []

        for shape in shapes:
            try:
                top_row = int(shape.TopLeftCell.Row)
                bottom_row = int(shape.BottomRightCell.Row)

                if bottom_row < start_row or top_row > end_row:
                    continue

                shape_type = int(shape.Type)
                if shape_type not in self.PICTURE_TYPES:
                    continue

                result["image_shape_count"] += 1

                if shape_type != self.LINKED_PICTURE_TYPE:
                    continue

                result["linked_picture_count"] += 1
                source_path = self._get_linked_source_path(shape)
                if not source_path:
                    continue

                normalized = str(Path(source_path))
                path_key = normalized.lower()
                if path_key in seen_paths:
                    continue

                seen_paths.add(path_key)
                result["linked_paths"].append(normalized)
                self._append_metadata_record(
                    result=result,
                    image_path=normalized,
                    path_source="excel_link",
                )
            except Exception:
                continue

        metadata_basenames = {
            Path(record["image_path"]).name.lower()
            for record in result["metadata_records"]
            if record.get("image_path")
        }

        for filename in dict.fromkeys(waveform_files or []):
            basename = Path(str(filename)).name
            if not basename or basename.lower() in metadata_basenames:
                continue

            match = self.find_original_png(basename)
            if match["match_count"] == 1:
                result["image_root_match_count"] += 1
                self._append_metadata_record(
                    result=result,
                    image_path=match["path"],
                    path_source="tek_remote_image",
                )
            elif match["match_count"] > 1:
                self.log(
                    f"[ImageTimeResolver] 동일 PNG 파일명 중복: "
                    f"{basename} / {match['match_count']}개"
                )

        result["metadata_time_list"] = [
            record["time"]
            for record in result["metadata_records"]
            if record.get("time")
        ]

        if result["image_shape_count"]:
            sheet_name = getattr(ws, "Name", "")
            self.log(
                f"[ImageTimeResolver] {sheet_name} rows {start_row}-{end_row}: "
                f"images={result['image_shape_count']}, "
                f"linked={result['linked_picture_count']}, "
                f"root_matches={result['image_root_match_count']}, "
                f"metadata={len(result['metadata_time_list'])}"
            )

        return result

    def find_original_png(self, filename):
        self._ensure_png_index()

        basename = Path(str(filename)).name.lower()
        matches = self._png_index.get(basename, []) if self._png_index else []

        return {
            "path": str(matches[0]) if len(matches) == 1 else "",
            "match_count": len(matches),
        }

    def _ensure_png_index(self):
        if self._png_index is not None:
            return

        self._png_index = {}

        if not self.image_root.is_dir():
            self.log(
                f"[ImageTimeResolver] 원본 PNG 폴더 없음: {self.image_root}"
            )
            return

        self.log(
            f"[ImageTimeResolver] 원본 PNG 인덱스 생성 시작: {self.image_root}"
        )

        file_count = 0
        for directory, _, filenames in os.walk(self.image_root):
            for filename in filenames:
                if not filename.lower().endswith(".png"):
                    continue

                full_path = Path(directory) / filename
                self._png_index.setdefault(filename.lower(), []).append(full_path)
                file_count += 1

        self.log(
            f"[ImageTimeResolver] 원본 PNG 인덱스 생성 완료: {file_count}개"
        )

    def _append_metadata_record(self, result, image_path, path_source):
        metadata = self.extract_png_metadata_time(image_path)
        if not metadata["time"]:
            return

        record_key = (
            str(image_path).lower(),
            metadata["time"],
            metadata["source"],
        )
        existing_keys = {
            (
                str(record.get("image_path", "")).lower(),
                record.get("time", ""),
                record.get("metadata_source", ""),
            )
            for record in result["metadata_records"]
        }
        if record_key in existing_keys:
            return

        result["metadata_records"].append(
            {
                "image_path": str(image_path),
                "path_source": path_source,
                "time": metadata["time"],
                "metadata_source": metadata["source"],
                "confidence": metadata["confidence"],
                "timezone": metadata["timezone"],
                "raw": metadata["raw"],
            }
        )

    @staticmethod
    def _get_linked_source_path(shape):
        candidates = []

        try:
            candidates.append(str(shape.LinkFormat.SourceFullName))
        except Exception:
            pass

        try:
            candidates.append(str(shape.Formula))
        except Exception:
            pass

        for candidate in candidates:
            text = candidate.strip().strip('"').strip("'")
            if not text:
                continue

            if text.lower().endswith(".png"):
                return text

            match = re.search(
                r"((?:[A-Za-z]:[\\/]|\\\\)[^\"']+?\.png)",
                text,
                re.IGNORECASE,
            )
            if match:
                return match.group(1)

        return ""

    def extract_png_metadata_time(self, image_path):
        result = {
            "time": "",
            "source": "",
            "confidence": "",
            "timezone": "",
            "raw": "",
        }

        path = Path(image_path)
        if not path.is_file() or path.suffix.lower() != ".png":
            return result

        candidates = []

        try:
            with Image.open(path) as image:
                exif = image.getexif()
                if exif:
                    for tag_id, value in exif.items():
                        tag_name = ExifTags.TAGS.get(tag_id, str(tag_id))
                        priority = self.EXIF_TIME_PRIORITY.get(tag_name)
                        if priority is None:
                            continue

                        parsed = self._parse_datetime_text(value)
                        if parsed:
                            candidates.append(
                                {
                                    "priority": priority,
                                    "time": parsed,
                                    "source": f"exif:{tag_name}",
                                    "confidence": "high" if priority == 10 else "medium",
                                    "timezone": "",
                                    "raw": str(value),
                                }
                            )

                text_items = {}
                text_items.update(getattr(image, "text", {}) or {})
                text_items.update(image.info or {})

                for key, value in text_items.items():
                    key_text = str(key).strip()
                    normalized_key = re.sub(
                        r"[^a-z0-9]", "", key_text.lower()
                    )
                    if normalized_key not in self.TEXT_TIME_KEYS:
                        continue

                    parsed = self._parse_datetime_text(value)
                    if not parsed:
                        continue

                    priority = 15 if normalized_key == "creationtime" else 25
                    candidates.append(
                        {
                            "priority": priority,
                            "time": parsed,
                            "source": f"png_text:{key_text}",
                            "confidence": "high" if priority == 15 else "medium",
                            "timezone": "",
                            "raw": str(value),
                        }
                    )
        except Exception:
            pass

        time_chunk = self._read_png_time_chunk(path)
        if time_chunk:
            candidates.append(
                {
                    "priority": 90,
                    "time": time_chunk,
                    "source": "png:tIME",
                    "confidence": "low",
                    "timezone": "UTC",
                    "raw": time_chunk.strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

        if not candidates:
            return result

        candidates.sort(key=lambda item: item["priority"])
        best = candidates[0]

        return {
            "time": best["time"].strftime("%Y-%m-%d %H:%M:%S"),
            "source": best["source"],
            "confidence": best["confidence"],
            "timezone": best["timezone"],
            "raw": best["raw"],
        }

    @staticmethod
    def _parse_datetime_text(value):
        if value is None:
            return None

        if isinstance(value, bytes):
            value = value.decode("utf-8", errors="ignore")

        text = str(value).replace("\x00", "").strip()
        if not text:
            return None

        try:
            return datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError:
            pass

        match = re.search(
            r"(\d{4})[-:/.](\d{1,2})[-:/.](\d{1,2})"
            r"[ T](\d{1,2}):(\d{2}):(\d{2})",
            text,
        )
        if match:
            try:
                return datetime(*map(int, match.groups()))
            except ValueError:
                return None

        for date_format in (
            "%Y%m%d%H%M%S",
            "%Y-%m-%d %H:%M:%S",
            "%Y/%m/%d %H:%M:%S",
            "%Y:%m:%d %H:%M:%S",
        ):
            try:
                return datetime.strptime(text, date_format)
            except ValueError:
                continue

        return None

    @staticmethod
    def _read_png_time_chunk(path):
        try:
            with Path(path).open("rb") as file:
                if file.read(8) != b"\x89PNG\r\n\x1a\n":
                    return None

                while True:
                    length_raw = file.read(4)
                    if len(length_raw) != 4:
                        return None

                    chunk_length = struct.unpack(">I", length_raw)[0]
                    chunk_type = file.read(4)
                    chunk_data = file.read(chunk_length)
                    file.read(4)  # CRC

                    if chunk_type == b"tIME" and len(chunk_data) == 7:
                        year, month, day, hour, minute, second = struct.unpack(
                            ">H5B", chunk_data
                        )
                        leap_second = second == 60
                        second = min(second, 59)
                        parsed = datetime(
                            year, month, day, hour, minute, second
                        )
                        if leap_second:
                            parsed += timedelta(seconds=1)
                        return parsed

                    if chunk_type == b"IEND":
                        return None
        except (OSError, ValueError, struct.error):
            return None
