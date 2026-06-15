import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
CONFIG_PATH = BASE_DIR / "config.json"

DEFAULT_CONFIG = {
    "target_extensions": [".log", ".txt"],
    "target_levels": ["WARNING", "ERROR", "CRITICAL"],
    "datetime_format": "%Y-%m-%d %H:%M:%S",
    "output_file": "log_analysis_report.xlsx",
    "category_keywords": {
        "database": ["database", "db", "sql", "mysql", "postgres"],
        "timeout": ["timeout", "timed out"],
        "permission": ["permission", "denied", "unauthorized", "forbidden"],
        "not_found": ["not found", "missing", "404"],
        "memory": ["memory", "out of memory", "oom"],
        "connection": ["connection", "connect", "network", "socket"],
        "api": ["api", "http", "request", "response"],
        "disk": ["disk", "storage", "no space"],
    },
}

SAMPLE_FILES = {
    "sample_app.log": """2026-06-15 09:00:01 INFO Application started
2026-06-15 09:01:12 INFO Loaded configuration file
2026-06-15 09:03:45 WARNING Disk usage is high: 82%
2026-06-15 09:05:10 ERROR Failed to connect database
2026-06-15 09:05:15 ERROR SQL query failed: missing column customer_id
2026-06-15 09:10:01 INFO Retry database connection
2026-06-15 09:10:09 ERROR Database connection timed out
2026-06-15 09:15:33 WARNING API response time is slow
2026-06-15 09:20:44 ERROR Permission denied while reading config file
2026-06-15 09:30:00 CRITICAL Out of memory while processing batch job
2026-06-15 09:35:22 INFO Batch job restarted
2026-06-15 09:40:18 ERROR User profile not found
BROKEN LOG LINE WITHOUT EXPECTED FORMAT
2026-06-15 09:50:11 WARNING Network connection is unstable
2026-06-15 10:00:00 INFO Application finished
""",
    "sample_server.log": """2026-06-15 10:00:01 INFO Server started on port 8080
2026-06-15 10:02:14 WARNING HTTP request took 3200 ms
2026-06-15 10:03:20 ERROR API request failed with status 500
2026-06-15 10:04:18 ERROR Connection refused by upstream server
2026-06-15 10:05:45 WARNING Disk free space is below threshold
2026-06-15 10:06:11 ERROR File not found: /var/data/orders.csv
2026-06-15 10:10:09 CRITICAL Service unavailable due to database outage
2026-06-15 10:12:00 INFO Health check passed
INVALID LINE: missing datetime and level
2026-06-15 10:15:32 ERROR Timeout while calling external API
""",
    "sample_legacy.txt": """2026-06-14 22:00:00 INFO Nightly job started
2026-06-14 22:01:05 WARNING Memory usage is high
2026-06-14 22:02:30 ERROR Failed to write output file: permission denied
2026-06-14 22:05:10 ERROR No space left on disk
2026-06-14 22:10:00 INFO Nightly job finished
""",
}

LOG_PATTERN = re.compile(
    r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\s+([A-Za-z]+)\s+(.*)$"
)

LOG_DETAIL_COLUMNS = [
    "file_name",
    "line_number",
    "datetime",
    "date",
    "level",
    "severity",
    "category",
    "message",
    "raw_line",
]
PARSE_ERROR_COLUMNS = ["file_name", "line_number", "raw_line", "error_message"]
FILE_SUMMARY_COLUMNS = [
    "file_name",
    "checked_lines",
    "target_logs",
    "warning_count",
    "error_count",
    "critical_count",
    "parse_error_count",
]


def create_sample_files():
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(
            json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    for file_name, content in SAMPLE_FILES.items():
        path = INPUT_DIR / file_name
        if not path.exists():
            path.write_text(content, encoding="utf-8")


def load_config():
    if not CONFIG_PATH.exists():
        create_sample_files()

    with CONFIG_PATH.open("r", encoding="utf-8") as file:
        return json.load(file)


def validate_config(config):
    required_keys = [
        "target_extensions",
        "target_levels",
        "datetime_format",
        "output_file",
        "category_keywords",
    ]
    for key in required_keys:
        if key not in config:
            raise ValueError(f"Missing required config key: {key}")

    target_extensions = config["target_extensions"]
    if not isinstance(target_extensions, list) or not target_extensions:
        raise ValueError("target_extensions must be a non-empty list.")
    normalized_extensions = []
    for extension in target_extensions:
        if not isinstance(extension, str) or not extension.strip():
            raise ValueError("Each target extension must be a non-empty string.")
        extension = extension.strip().lower()
        if not extension.startswith("."):
            raise ValueError(f"Target extension must start with '.': {extension}")
        normalized_extensions.append(extension)

    target_levels = config["target_levels"]
    if not isinstance(target_levels, list) or not target_levels:
        raise ValueError("target_levels must be a non-empty list.")
    normalized_levels = []
    for level in target_levels:
        if not isinstance(level, str) or not level.strip():
            raise ValueError("Each target level must be a non-empty string.")
        normalized_levels.append(level.strip().upper())

    datetime_format = config["datetime_format"]
    if not isinstance(datetime_format, str) or not datetime_format.strip():
        raise ValueError("datetime_format must be a non-empty string.")

    output_file = config["output_file"]
    if (
        not isinstance(output_file, str)
        or not output_file.strip()
        or not output_file.lower().endswith(".xlsx")
    ):
        raise ValueError("output_file must be a non-empty .xlsx file name.")

    category_keywords = config["category_keywords"]
    if not isinstance(category_keywords, dict):
        raise ValueError("category_keywords must be a dictionary.")
    normalized_keywords = {}
    for category, keywords in category_keywords.items():
        if not isinstance(category, str) or not category.strip():
            raise ValueError("Each category name must be a non-empty string.")
        if not isinstance(keywords, list):
            raise ValueError(f"Keywords for {category} must be a list.")
        normalized_list = []
        for keyword in keywords:
            if not isinstance(keyword, str) or not keyword.strip():
                raise ValueError(f"Each keyword for {category} must be non-empty.")
            normalized_list.append(keyword.strip())
        normalized_keywords[category.strip()] = normalized_list

    validated = dict(config)
    validated["target_extensions"] = normalized_extensions
    validated["target_levels"] = normalized_levels
    validated["datetime_format"] = datetime_format.strip()
    validated["output_file"] = output_file.strip()
    validated["category_keywords"] = normalized_keywords
    return validated


def find_log_files(config):
    if not INPUT_DIR.exists():
        return []
    target_extensions = set(config["target_extensions"])
    return sorted(
        [
            path
            for path in INPUT_DIR.iterdir()
            if path.is_file() and path.suffix.lower() in target_extensions
        ],
        key=lambda path: path.name.lower(),
    )


def parse_log_line(line, datetime_format):
    raw_line = line.rstrip("\n\r")
    if not raw_line.strip():
        return {"status": "empty", "raw_line": raw_line}

    match = LOG_PATTERN.match(raw_line)
    if not match:
        return {
            "status": "parse_error",
            "raw_line": raw_line,
            "error_message": "Line does not match expected log format",
        }

    datetime_text, level, message = match.groups()
    try:
        datetime.strptime(datetime_text, datetime_format)
    except ValueError as exc:
        return {
            "status": "parse_error",
            "raw_line": raw_line,
            "error_message": f"Invalid datetime: {exc}",
        }

    return {
        "status": "ok",
        "datetime": datetime_text,
        "level": level.upper(),
        "message": message,
        "raw_line": raw_line,
    }


def classify_log_category(message, category_keywords):
    lowered_message = message.lower()
    for category, keywords in category_keywords.items():
        for keyword in keywords:
            if keyword.lower() in lowered_message:
                return category
    return "other"


def get_severity(level):
    severity_map = {
        "CRITICAL": "high",
        "ERROR": "high",
        "WARNING": "medium",
    }
    return severity_map.get(level, "low")


def make_empty_file_stats(file_name):
    return {
        "file_name": file_name,
        "checked_lines": 0,
        "target_logs": 0,
        "warning_count": 0,
        "error_count": 0,
        "critical_count": 0,
        "parse_error_count": 0,
    }


def read_text_lines(path):
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp932"):
        try:
            with path.open("r", encoding=encoding) as file:
                return file.readlines()
        except UnicodeDecodeError as exc:
            last_error = exc
    raise last_error


def analyze_log_file(path, config):
    log_records = []
    parse_errors = []
    file_stats = make_empty_file_stats(path.name)
    target_levels = set(config["target_levels"])

    try:
        lines = read_text_lines(path)
    except Exception as exc:
        parse_errors.append(
            {
                "file_name": path.name,
                "line_number": "",
                "raw_line": "",
                "error_message": f"failed to read file: {exc}",
            }
        )
        file_stats["parse_error_count"] = 1
        return log_records, parse_errors, file_stats

    for line_number, line in enumerate(lines, start=1):
        parsed = parse_log_line(line, config["datetime_format"])
        if parsed["status"] == "empty":
            continue

        file_stats["checked_lines"] += 1

        if parsed["status"] == "parse_error":
            parse_errors.append(
                {
                    "file_name": path.name,
                    "line_number": line_number,
                    "raw_line": parsed["raw_line"],
                    "error_message": parsed["error_message"],
                }
            )
            file_stats["parse_error_count"] += 1
            continue

        level = parsed["level"]
        if level not in target_levels:
            continue

        category = classify_log_category(
            parsed["message"], config["category_keywords"]
        )
        record = {
            "file_name": path.name,
            "line_number": line_number,
            "datetime": parsed["datetime"],
            "date": parsed["datetime"][:10],
            "level": level,
            "severity": get_severity(level),
            "category": category,
            "message": parsed["message"],
            "raw_line": parsed["raw_line"],
        }
        log_records.append(record)

        file_stats["target_logs"] += 1
        if level == "WARNING":
            file_stats["warning_count"] += 1
        elif level == "ERROR":
            file_stats["error_count"] += 1
        elif level == "CRITICAL":
            file_stats["critical_count"] += 1

    return log_records, parse_errors, file_stats


def analyze_log_files(log_files, config):
    result = {
        "log_records": [],
        "parse_errors": [],
        "file_stats": [],
        "checked_files": len(log_files),
        "checked_lines": 0,
    }

    for path in log_files:
        records, errors, stats = analyze_log_file(path, config)
        result["log_records"].extend(records)
        result["parse_errors"].extend(errors)
        result["file_stats"].append(stats)
        result["checked_lines"] += stats["checked_lines"]

    return result


def make_summary_df(result, started_at, finished_at):
    records = result["log_records"]
    parse_errors = result["parse_errors"]
    rows = [
        ("checked_files", result["checked_files"]),
        ("checked_lines", result["checked_lines"]),
        ("target_logs", len(records)),
        ("warning_count", sum(1 for row in records if row["level"] == "WARNING")),
        ("error_count", sum(1 for row in records if row["level"] == "ERROR")),
        ("critical_count", sum(1 for row in records if row["level"] == "CRITICAL")),
        ("parse_error_count", len(parse_errors)),
        ("started_at", started_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("finished_at", finished_at.strftime("%Y-%m-%d %H:%M:%S")),
    ]
    return pd.DataFrame(rows, columns=["item", "value"])


def make_log_details_df(log_records):
    df = pd.DataFrame(log_records, columns=LOG_DETAIL_COLUMNS)
    if not df.empty:
        df = df.sort_values(
            by=["datetime", "file_name", "line_number"], kind="stable"
        )
    return df


def make_level_summary_df(log_records, target_levels):
    counts = {level: 0 for level in target_levels}
    for record in log_records:
        counts[record["level"]] = counts.get(record["level"], 0) + 1

    order = ["CRITICAL", "ERROR", "WARNING"]
    ordered_levels = [level for level in order if level in counts]
    ordered_levels.extend(level for level in counts if level not in ordered_levels)

    return pd.DataFrame(
        [{"level": level, "count": counts[level]} for level in ordered_levels],
        columns=["level", "count"],
    )


def make_category_summary_df(log_records):
    columns = ["category", "count", "warning_count", "error_count", "critical_count"]
    if not log_records:
        return pd.DataFrame(columns=columns)

    df = pd.DataFrame(log_records)
    grouped = (
        df.groupby("category")
        .agg(
            count=("category", "size"),
            warning_count=("level", lambda values: (values == "WARNING").sum()),
            error_count=("level", lambda values: (values == "ERROR").sum()),
            critical_count=("level", lambda values: (values == "CRITICAL").sum()),
        )
        .reset_index()
        .sort_values(by=["count", "category"], ascending=[False, True], kind="stable")
    )
    return grouped[columns]


def make_file_summary_df(file_stats):
    df = pd.DataFrame(file_stats, columns=FILE_SUMMARY_COLUMNS)
    if not df.empty:
        df = df.sort_values(by="file_name", kind="stable")
    return df


def make_timeline_summary_df(log_records):
    columns = ["date", "warning_count", "error_count", "critical_count", "total_count"]
    if not log_records:
        return pd.DataFrame(columns=columns)

    df = pd.DataFrame(log_records)
    grouped = (
        df.groupby("date")
        .agg(
            warning_count=("level", lambda values: (values == "WARNING").sum()),
            error_count=("level", lambda values: (values == "ERROR").sum()),
            critical_count=("level", lambda values: (values == "CRITICAL").sum()),
            total_count=("level", "size"),
        )
        .reset_index()
        .sort_values(by="date", kind="stable")
    )
    return grouped[columns]


def make_parse_errors_df(parse_errors):
    df = pd.DataFrame(parse_errors, columns=PARSE_ERROR_COLUMNS)
    if not df.empty:
        df = df.sort_values(by=["file_name", "line_number"], kind="stable")
    return df


def make_config_df(config):
    rows = []
    for key, value in config.items():
        if isinstance(value, list):
            display_value = ", ".join(value)
        elif isinstance(value, dict):
            display_value = json.dumps(value, ensure_ascii=False)
        else:
            display_value = str(value)
        rows.append({"key": key, "value": display_value})
    return pd.DataFrame(rows, columns=["key", "value"])


def make_output_path(config):
    OUTPUT_DIR.mkdir(exist_ok=True)
    return OUTPUT_DIR / config["output_file"]


def adjust_column_width(worksheet, df):
    for column_index, column_name in enumerate(df.columns):
        values = [str(column_name)]
        if not df.empty:
            values.extend(str(value) for value in df[column_name].fillna("").tolist())
        width = min(max(len(value) for value in values) + 2, 80)
        worksheet.set_column(column_index, column_index, width)


def format_worksheet(workbook, worksheet, df):
    header_format = workbook.add_format(
        {"bold": True, "bg_color": "#D9EAF7", "border": 1}
    )
    for column_index, column_name in enumerate(df.columns):
        worksheet.write(0, column_index, column_name, header_format)

    row_count = max(len(df), 1)
    column_count = max(len(df.columns) - 1, 0)
    worksheet.autofilter(0, 0, row_count, column_count)
    worksheet.freeze_panes(1, 0)
    adjust_column_width(worksheet, df)


def save_excel_report(result, config, started_at, finished_at):
    output_path = make_output_path(config)
    sheets = {
        "summary": make_summary_df(result, started_at, finished_at),
        "log_details": make_log_details_df(result["log_records"]),
        "level_summary": make_level_summary_df(
            result["log_records"], config["target_levels"]
        ),
        "category_summary": make_category_summary_df(result["log_records"]),
        "file_summary": make_file_summary_df(result["file_stats"]),
        "timeline_summary": make_timeline_summary_df(result["log_records"]),
        "parse_errors": make_parse_errors_df(result["parse_errors"]),
        "config": make_config_df(config),
    }

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            format_worksheet(writer.book, worksheet, df)

    return output_path


def main():
    started_at = datetime.now()
    create_sample_files()
    config = validate_config(load_config())
    log_files = find_log_files(config)

    if not log_files:
        print("No log files found in input directory.")

    result = analyze_log_files(log_files, config)
    finished_at = datetime.now()
    output_path = save_excel_report(result, config, started_at, finished_at)

    print(f"Log files found: {len(log_files)}")
    print(f"Checked lines: {result['checked_lines']}")
    print(f"Target logs found: {len(result['log_records'])}")
    print(f"Parse errors: {len(result['parse_errors'])}")
    print(f"Excel saved: {output_path.relative_to(BASE_DIR)}")


if __name__ == "__main__":
    main()
