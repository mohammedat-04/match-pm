#!/usr/bin/env python3
"""Create a Numbers-style spreadsheet from project vision_results.json files.

By default the workbook follows the same table layout as Untitled.numbers.
Apple Numbers can open the generated .xlsx directly. On macOS, pass --numbers
to ask the Numbers app to save a native .numbers file as well.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import re
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape


FLOAT_RE = r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?"
DATE_SUFFIX_RE = re.compile(r"_(\d{8})_(\d{2})_(\d{2})_(\d{2})$")
CIRCLE_RE = re.compile(
    rf"VisionCircle\(center_point=.*?VisionPoint\("
    rf"axis_value_1=({FLOAT_RE}), axis_value_2=({FLOAT_RE}), "
    rf"axis_suffix_1='([^']*)', axis_suffix_2='([^']*)'\), "
    rf"radius=({FLOAT_RE})\)"
)
POINTS_BLOCK_RE = re.compile(r"points=\[(.*?)\], lines=", re.DOTALL)
POINT_RE = re.compile(
    rf"VisionPoint\(axis_value_1=({FLOAT_RE}), axis_value_2=({FLOAT_RE}), "
    rf"axis_suffix_1='([^']*)', axis_suffix_2='([^']*)'\)"
)


@dataclass
class ParsedResult:
    path: Path
    group: str
    run_name: str
    target: str
    provider: str
    model: str
    run_timestamp: str
    exec_timestamp: str
    vision_ok: bool | None
    result_type: str
    detected: bool
    succeeded: bool
    point_x: float | None
    point_y: float | None
    circle_x: float | None
    circle_y: float | None
    circle_r: float | None
    circle_x_px: float | None
    circle_y_px: float | None
    circle_r_px: float | None
    debug: str


@dataclass(frozen=True)
class Calibration:
    pixel_scale: float
    pixel_origin_x: float
    pixel_origin_y: float
    reference_x_px: float
    reference_y_px: float
    reference_r_px: float
    center_tolerance_px: float
    radius_tolerance_px: float


def main() -> int:
    args = parse_args()
    root = args.root.resolve()
    output = args.output.resolve()

    if output.exists() and not args.force:
        raise SystemExit(f"{output} already exists. Use --force to overwrite it.")

    calibration = Calibration(
        pixel_scale=args.pixel_scale,
        pixel_origin_x=args.pixel_origin_x,
        pixel_origin_y=args.pixel_origin_y,
        reference_x_px=args.reference_x_px,
        reference_y_px=args.reference_y_px,
        reference_r_px=args.reference_r_px,
        center_tolerance_px=args.center_tolerance_px,
        radius_tolerance_px=args.radius_tolerance_px,
    )

    results = collect_results(root, calibration)
    if not results:
        raise SystemExit(f"No vision_results.json files found under {root}")

    if args.mode == "untitled":
        results = add_missing_untitled_runs(root, results)
        sheets = build_untitled_sheets(results)
        result_count = count_untitled_results(results)
    else:
        sheets = build_sheets(results, include_group_sheets=not args.no_group_sheets)
        result_count = len(results)

    write_xlsx(output, sheets)
    print(f"Wrote {output} ({result_count} results, {len(sheets)} sheets)")

    if args.numbers is not None:
        numbers_output = args.numbers.resolve()
        export_numbers(output, numbers_output, force=args.force)
        print(f"Wrote {numbers_output}")

    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build an Excel/Numbers-ready workbook from vision_results.json files."
    )
    parser.add_argument("--root", type=Path, default=Path("."), help="folder to scan")
    parser.add_argument(
        "--mode",
        choices=("untitled", "all"),
        default="untitled",
        help="untitled matches Untitled.numbers; all writes every parsed result",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("vision_results.xlsx"),
        help="xlsx workbook to write",
    )
    parser.add_argument(
        "--numbers",
        nargs="?",
        const=Path("vision_results.numbers"),
        type=Path,
        help="also export a native .numbers file using the macOS Numbers app",
    )
    parser.add_argument("--force", action="store_true", help="overwrite output files")
    parser.add_argument(
        "--no-group-sheets",
        action="store_true",
        help="with --mode all, write only the all-results and success-rate sheets",
    )
    parser.add_argument(
        "--pixel-scale",
        type=float,
        default=4.39561,
        help="metric-to-pixel scale used for vacuum gripper circles",
    )
    parser.add_argument("--pixel-origin-x", type=float, default=1296.0)
    parser.add_argument("--pixel-origin-y", type=float, default=972.0)
    parser.add_argument("--reference-x-px", type=float, default=1294.0)
    parser.add_argument("--reference-y-px", type=float, default=1038.0)
    parser.add_argument("--reference-r-px", type=float, default=221.0)
    parser.add_argument("--center-tolerance-px", type=float, default=10.0)
    parser.add_argument("--radius-tolerance-px", type=float, default=10.0)
    return parser.parse_args()


def collect_results(root: Path, calibration: Calibration) -> list[ParsedResult]:
    rows: list[ParsedResult] = []
    for json_path in sorted(root.rglob("vision_results.json")):
        if ".git" in json_path.parts:
            continue
        rows.append(parse_result(root, json_path, calibration))
    return rows


def add_missing_untitled_runs(root: Path, results: list[ParsedResult]) -> list[ParsedResult]:
    rows = list(results)
    known_parents = {row.path.parent for row in rows}

    for spec in untitled_table_specs():
        for group in spec["source_groups"]:
            group_dir = root / group
            if not group_dir.is_dir():
                continue
            for run_dir in sorted(path for path in group_dir.iterdir() if path.is_dir()):
                rel_parent = Path(group) / run_dir.name
                if rel_parent in known_parents:
                    continue

                metadata = parse_run_metadata(run_dir.name)
                if metadata["target"] != "vacuum gripper" or metadata["model"] != spec["model"]:
                    continue

                rows.append(missing_result(group, run_dir.name, metadata))
                known_parents.add(rel_parent)

    return rows


def missing_result(group: str, run_name: str, metadata: dict[str, str]) -> ParsedResult:
    return ParsedResult(
        path=Path(group) / run_name / "vision_results.json",
        group=group,
        run_name=run_name,
        target=metadata["target"],
        provider=metadata["provider"],
        model=metadata["model"],
        run_timestamp=metadata["run_timestamp"],
        exec_timestamp="",
        vision_ok=False,
        result_type="none",
        detected=False,
        succeeded=False,
        point_x=None,
        point_y=None,
        circle_x=None,
        circle_y=None,
        circle_r=None,
        circle_x_px=None,
        circle_y_px=None,
        circle_r_px=None,
        debug="missing vision_results.json",
    )


def parse_result(root: Path, json_path: Path, calibration: Calibration) -> ParsedResult:
    data = json.loads(json_path.read_text(encoding="utf-8"))
    text = str(data.get("vision_result", ""))
    debug = "; ".join(str(item) for item in data.get("process_debug", []))

    rel_path = json_path.relative_to(root)
    group = rel_path.parts[0] if len(rel_path.parts) > 1 else ""
    run_name = json_path.parent.name
    metadata = parse_run_metadata(run_name)

    circles = parse_circles(text)
    points = parse_points(text)

    circle = circles[0] if circles else None
    point = points[0] if points else None
    vision_ok = parse_bool_field(text, "vision_ok")
    exec_timestamp = parse_string_field(text, "exec_timestamp")
    detected = bool(circles or points)

    circle_x = circle["x"] if circle else None
    circle_y = circle["y"] if circle else None
    circle_r = circle["r"] if circle else None
    circle_x_px = circle_y_px = circle_r_px = None
    if circle is not None:
        circle_x_px = calibration.pixel_origin_x - circle_x / calibration.pixel_scale
        circle_y_px = calibration.pixel_origin_y + circle_y / calibration.pixel_scale
        circle_r_px = circle_r / calibration.pixel_scale

    if circle is not None:
        succeeded = circle_succeeded(circle_x_px, circle_y_px, circle_r_px, calibration)
        result_type = "circle"
    elif point is not None:
        succeeded = bool(vision_ok)
        result_type = "point"
    else:
        succeeded = False
        result_type = "none"

    return ParsedResult(
        path=rel_path,
        group=group,
        run_name=run_name,
        target=metadata["target"],
        provider=metadata["provider"],
        model=metadata["model"],
        run_timestamp=metadata["run_timestamp"],
        exec_timestamp=exec_timestamp,
        vision_ok=vision_ok,
        result_type=result_type,
        detected=detected,
        succeeded=succeeded,
        point_x=point["x"] if point else None,
        point_y=point["y"] if point else None,
        circle_x=circle_x,
        circle_y=circle_y,
        circle_r=circle_r,
        circle_x_px=circle_x_px,
        circle_y_px=circle_y_px,
        circle_r_px=circle_r_px,
        debug=debug,
    )


def parse_run_metadata(run_name: str) -> dict[str, str]:
    run_timestamp = ""
    name_without_date = run_name
    match = DATE_SUFFIX_RE.search(run_name)
    if match:
        date_part, hour, minute, second = match.groups()
        run_timestamp = (
            f"{date_part[0:4]}-{date_part[4:6]}-{date_part[6:8]} "
            f"{hour}:{minute}:{second}"
        )
        name_without_date = run_name[: match.start()]

    if "_huggingface_" not in name_without_date:
        return {
            "target": name_without_date.replace("_", " "),
            "provider": "",
            "model": "",
            "run_timestamp": run_timestamp,
        }

    target_slug, provider_and_model = name_without_date.split("_huggingface_", 1)
    provider, _, model_slug = provider_and_model.partition("_")
    return {
        "target": target_slug.replace("_", " "),
        "provider": provider.replace("_", " "),
        "model": model_slug.replace("_", " "),
        "run_timestamp": run_timestamp,
    }


def parse_circles(text: str) -> list[dict[str, float]]:
    circles: list[dict[str, float]] = []
    for match in CIRCLE_RE.finditer(text):
        x, y, _suffix_x, _suffix_y, radius = match.groups()
        circles.append({"x": float(x), "y": float(y), "r": float(radius)})
    return circles


def parse_points(text: str) -> list[dict[str, float]]:
    block_match = POINTS_BLOCK_RE.search(text)
    if not block_match:
        return []

    points: list[dict[str, float]] = []
    for match in POINT_RE.finditer(block_match.group(1)):
        x, y, _suffix_x, _suffix_y = match.groups()
        points.append({"x": float(x), "y": float(y)})
    return points


def parse_bool_field(text: str, name: str) -> bool | None:
    match = re.search(rf"\b{name}=(True|False)", text)
    if not match:
        return None
    return match.group(1) == "True"


def parse_string_field(text: str, name: str) -> str:
    match = re.search(rf"\b{name}='([^']*)'", text)
    return match.group(1) if match else ""


def circle_succeeded(
    x_px: float | None, y_px: float | None, r_px: float | None, calibration: Calibration
) -> bool:
    if x_px is None or y_px is None or r_px is None:
        return False
    center_distance = math.hypot(
        x_px - calibration.reference_x_px, y_px - calibration.reference_y_px
    )
    radius_delta = abs(r_px - calibration.reference_r_px)
    return (
        center_distance <= calibration.center_tolerance_px
        and radius_delta <= calibration.radius_tolerance_px
    )


def build_untitled_sheets(results: list[ParsedResult]) -> list[tuple[str, list[list[Any]]]]:
    specs = untitled_table_specs()
    sheets: list[tuple[str, list[list[Any]]]] = [
        ("Export Summary", build_untitled_export_summary_sheet()),
        ("Vision Results - Vacuum gripper", build_untitled_hough_sheet()),
    ]

    table_results: list[tuple[dict[str, Any], list[ParsedResult]]] = []
    for spec in specs:
        selected = select_untitled_results(results, spec)
        table_results.append((spec, selected))
        sheets.append((spec["sheet_name"], build_untitled_result_sheet(spec, selected)))

    sheets.append(("Vision Results - Drawings", []))
    sheets.append(("Sheet 1 - Success Rate", build_untitled_success_rate_sheet(table_results)))
    return sheets


def count_untitled_results(results: list[ParsedResult]) -> int:
    return sum(len(select_untitled_results(results, spec)) for spec in untitled_table_specs())


def untitled_table_specs() -> list[dict[str, Any]]:
    return old_untitled_table_specs() + prompt_table_specs()


def old_untitled_table_specs() -> list[dict[str, Any]]:
    return [
        {
            "sheet_name": "Vision Results - Full - Qwen Qw",
            "title": "Full - Qwen Qwen3.6-35B-A3B",
            "type": "Full",
            "model": "Qwen Qwen3.6-35B-A3B",
            "source_groups": {"Full", "Full specs"},
            "display_group": "Full",
        },
        {
            "sheet_name": "Vision Results - Full - zai-org",
            "title": "Full - zai-org GLM-4.5V",
            "type": "Full",
            "model": "zai-org GLM-4.5V",
            "source_groups": {"Full", "Full specs"},
            "display_group": "Full",
        },
        {
            "sheet_name": "Vision Results - Namee only - Q",
            "title": "Namee only - Qwen Qwen3.6-35B-A3B",
            "type": "Namee only",
            "model": "Qwen Qwen3.6-35B-A3B",
            "source_groups": {"Name only", "Namee only"},
            "display_group": "Namee only",
        },
        {
            "sheet_name": "Vision Results - Namee only - z",
            "title": "Namee only - zai-org GLM-4.5V",
            "type": "Namee only",
            "model": "zai-org GLM-4.5V",
            "source_groups": {"Name only", "Namee only"},
            "display_group": "Namee only",
        },
    ]


def prompt_table_specs() -> list[dict[str, Any]]:
    return [
        prompt_table_spec(
            "new prompt full specs",
            "Qwen Qwen3.6-35B-A3B",
            {"new promt full specs", "new prompt full specs"},
            "New prompt full - Qwen",
        ),
        prompt_table_spec(
            "new prompt full specs",
            "google gemma-4-31B-it",
            {"new promt full specs", "new prompt full specs"},
            "New prompt full - gemma",
        ),
        prompt_table_spec(
            "new prompt full specs",
            "moonshotai Kimi-K2.5",
            {"new promt full specs", "new prompt full specs"},
            "New prompt full - Kimi",
        ),
        prompt_table_spec(
            "new prompt full specs",
            "zai-org GLM-4.5V",
            {"new promt full specs", "new prompt full specs"},
            "New prompt full - GLM",
        ),
        prompt_table_spec(
            "new prompt 1full specs",
            "google gemma-4-31B-it",
            {"new promt 1full specs", "new prompt 1full specs"},
            "New prompt 1full - gemma",
        ),
        prompt_table_spec(
            "new prompt 1full specs",
            "MiniMaxAI MiniMax-M3",
            {"new promt 1full specs", "new prompt 1full specs"},
            "New prompt 1full - MiniMax",
        ),
        prompt_table_spec(
            "new prompt name only",
            "Qwen Qwen3.6-35B-A3B",
            {"new promt name only", "new prompt name only"},
            "New prompt name - Qwen",
        ),
        prompt_table_spec(
            "new prompt name only",
            "google gemma-4-31B-it",
            {"new promt name only", "new prompt name only"},
            "New prompt name - gemma",
        ),
        prompt_table_spec(
            "new prompt name only",
            "MiniMaxAI MiniMax-M3",
            {"new promt name only", "new prompt name only"},
            "New prompt name - MiniMax",
        ),
        prompt_table_spec(
            "new prompt1 name only",
            "google gemma-4-31B-it",
            {"new promt1 name only", "new prompt1 name only"},
            "New prompt1 name - gemma",
        ),
        prompt_table_spec(
            "new prompt1 name only",
            "MiniMaxAI MiniMax-M3",
            {"new promt1 name only", "new prompt1 name only"},
            "New prompt1 name - MiniMax",
        ),
    ]


def prompt_table_spec(
    group_label: str, model: str, source_groups: set[str], sheet_name: str
) -> dict[str, Any]:
    return {
        "sheet_name": sheet_name,
        "title": f"{group_label} - {model}",
        "type": group_label,
        "model": model,
        "source_groups": source_groups,
        "display_group": group_label,
    }


def select_untitled_results(
    results: list[ParsedResult], spec: dict[str, Any]
) -> list[ParsedResult]:
    selected = [
        result
        for result in results
        if result.group in spec["source_groups"]
        and result.target == "vacuum gripper"
        and result.model == spec["model"]
    ]
    return sorted(selected, key=lambda item: (item.run_timestamp, str(item.path)))


def build_untitled_export_summary_sheet() -> list[list[Any]]:
    rows: list[list[Any]] = [
        [
            "This document was exported from Numbers. Each table was converted to an Excel worksheet. "
            "All other objects on each Numbers sheet were placed on separate worksheets. "
            "Please be aware that formula calculations may differ in Excel.",
            "",
            "",
        ],
        ["Numbers Sheet Name", "Numbers Table Name", "Excel Worksheet Name"],
        ["Vision Results", "", ""],
        ["", "Vacuum gripper coax OpenCV HoughCircles", "Vision Results - Vacuum gripper"],
    ]
    for spec in untitled_table_specs():
        rows.append(["", spec["title"], spec["sheet_name"]])
    rows.extend(
        [
            ["", '"All Drawings from the Sheet"', "Vision Results - Drawings"],
            ["Sheet 1", "", ""],
            ["", "Success Rate", "Sheet 1 - Success Rate"],
        ]
    )
    return rows


def build_untitled_hough_sheet() -> list[list[Any]]:
    return [
        ["Vacuum gripper coax OpenCV HoughCircles", "", "", ""],
        ["", "x", "y", "r"],
        ["HoughCircles", 1294, 1038, 221],
        ["Cam_CS", 8.79, 290.11, 971.43],
    ]


def build_untitled_result_sheet(
    spec: dict[str, Any], results: list[ParsedResult]
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        [spec["title"], "", "", "", "", "", ""],
        ["File name", "x", "y", "r", "x_px", "y_px", "r_px"],
    ]
    for result in results:
        rows.append(
            [
                f"{spec['display_group']}/{result.run_name}",
                legacy_metric(result.circle_x),
                legacy_metric(result.circle_y),
                legacy_metric(result.circle_r),
                legacy_pixel(result.circle_x_px),
                legacy_pixel(result.circle_y_px),
                legacy_pixel(result.circle_r_px),
            ]
        )
    return rows


def build_untitled_success_rate_sheet(
    table_results: list[tuple[dict[str, Any], list[ParsedResult]]]
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        ["Success Rate", "", "", "", "", ""],
        ["Type", "Model", "Total", "Detected", "Succeeded", "Success rate"],
    ]
    overall_total = 0
    overall_detected = 0
    overall_succeeded = 0

    for spec, results in table_results:
        total = len(results)
        detected = sum(result.circle_r is not None for result in results)
        succeeded = sum(result.succeeded for result in results)
        overall_total += total
        overall_detected += detected
        overall_succeeded += succeeded
        rows.append(
            [
                spec["type"],
                spec["model"],
                total,
                detected,
                succeeded,
                percent(succeeded, total),
            ]
        )

    rows.append(
        [
            "Overall",
            "All",
            overall_total,
            overall_detected,
            overall_succeeded,
            percent(overall_succeeded, overall_total),
        ]
    )
    return rows


def legacy_metric(value: float | None) -> float | str:
    if value is None:
        return "-"
    return value


def legacy_pixel(value: float | None) -> float | str:
    if value is None:
        return "-"
    return round(value, 3)


def build_sheets(
    results: list[ParsedResult], include_group_sheets: bool
) -> list[tuple[str, list[list[Any]]]]:
    sheets: list[tuple[str, list[list[Any]]]] = [
        ("All Results", build_all_results_sheet(results)),
        ("Success Rate", build_success_rate_sheet(results)),
    ]

    if include_group_sheets:
        grouped: dict[tuple[str, str, str, str], list[ParsedResult]] = {}
        for row in results:
            key = (row.group, row.target, row.provider, row.model)
            grouped.setdefault(key, []).append(row)

        for key in sorted(grouped):
            group, target, provider, model = key
            title_parts = [part for part in (group, target, provider, model) if part]
            title = " - ".join(title_parts) or "Results"
            sheet_name = compact_group_sheet_name(group, provider, model)
            sheets.append((sheet_name, build_group_sheet(grouped[key], title)))

    return sheets


def build_all_results_sheet(results: list[ParsedResult]) -> list[list[Any]]:
    rows: list[list[Any]] = [
        ["All Vision Results"],
        [
            "File name",
            "Group",
            "Target",
            "Provider",
            "Model",
            "Run timestamp",
            "Exec timestamp",
            "Vision OK",
            "Result type",
            "Detected",
            "Succeeded",
            "point_x",
            "point_y",
            "circle_x",
            "circle_y",
            "circle_r",
            "circle_x_px",
            "circle_y_px",
            "circle_r_px",
            "Debug",
        ],
    ]
    for result in results:
        rows.append(
            [
                str(result.path.parent),
                result.group,
                result.target,
                result.provider,
                result.model,
                result.run_timestamp,
                result.exec_timestamp,
                bool_label(result.vision_ok),
                result.result_type,
                bool_label(result.detected),
                bool_label(result.succeeded),
                format_number(result.point_x),
                format_number(result.point_y),
                format_number(result.circle_x),
                format_number(result.circle_y),
                format_number(result.circle_r),
                format_number(result.circle_x_px),
                format_number(result.circle_y_px),
                format_number(result.circle_r_px),
                result.debug,
            ]
        )
    return rows


def build_success_rate_sheet(results: list[ParsedResult]) -> list[list[Any]]:
    summary: dict[tuple[str, str, str, str], dict[str, int]] = {}
    for result in results:
        key = (result.group, result.target, result.provider, result.model)
        counts = summary.setdefault(key, {"total": 0, "detected": 0, "succeeded": 0})
        counts["total"] += 1
        counts["detected"] += int(result.detected)
        counts["succeeded"] += int(result.succeeded)

    rows: list[list[Any]] = [
        ["Success Rate"],
        ["Group", "Target", "Provider", "Model", "Total", "Detected", "Succeeded", "Success rate"],
    ]
    overall = {"total": 0, "detected": 0, "succeeded": 0}
    for key in sorted(summary):
        counts = summary[key]
        overall["total"] += counts["total"]
        overall["detected"] += counts["detected"]
        overall["succeeded"] += counts["succeeded"]
        rows.append(
            [
                *key,
                counts["total"],
                counts["detected"],
                counts["succeeded"],
                percent(counts["succeeded"], counts["total"]),
            ]
        )

    rows.append([])
    rows.append(
        [
            "Overall",
            "All",
            "",
            "",
            overall["total"],
            overall["detected"],
            overall["succeeded"],
            percent(overall["succeeded"], overall["total"]),
        ]
    )
    return rows


def compact_group_sheet_name(group: str, provider: str, model: str) -> str:
    parts = [compact_group(group), compact_provider(provider), compact_model(model)]
    return " ".join(part for part in parts if part) or "Results"


def compact_group(group: str) -> str:
    normalized = group.replace("_", " ").strip().lower()
    replacements = {
        "corner full": "Corner",
        "corner name only": "CornerName",
        "full specs": "Full",
        "name only": "Name",
        "namee only": "Namee",
        "new promt full specs": "NewPromptFull",
        "new prompt full specs": "NewPromptFull",
        "new promt 1full specs": "NewPrompt1Full",
        "new prompt 1full specs": "NewPrompt1Full",
        "new promt name only": "NewPromptName",
        "new prompt name only": "NewPromptName",
        "new promt1 name only": "NewPrompt1Name",
        "new prompt1 name only": "NewPrompt1Name",
    }
    if normalized in replacements:
        return replacements[normalized]
    return "".join(token[:4].title() for token in normalized.split())[:10]


def compact_provider(provider: str) -> str:
    provider = provider.strip()
    if not provider:
        return ""
    return provider.replace("-", "")[:5]


def compact_model(model: str) -> str:
    tokens = model.strip().split()
    if not tokens:
        return ""
    tail = " ".join(tokens[1:] if len(tokens) > 1 else tokens)
    compact = re.sub(r"[^A-Za-z0-9.]+", "", tail)
    return compact[:14] if compact else tokens[0][:14]


def build_group_sheet(results: list[ParsedResult], title: str) -> list[list[Any]]:
    rows: list[list[Any]] = [
        [title],
        [
            "File name",
            "Run timestamp",
            "Exec timestamp",
            "Vision OK",
            "Result type",
            "Detected",
            "Succeeded",
            "point_x",
            "point_y",
            "circle_x",
            "circle_y",
            "circle_r",
            "circle_x_px",
            "circle_y_px",
            "circle_r_px",
            "Debug",
        ],
    ]
    for result in sorted(results, key=lambda item: (item.run_timestamp, str(item.path))):
        rows.append(
            [
                str(result.path.parent),
                result.run_timestamp,
                result.exec_timestamp,
                bool_label(result.vision_ok),
                result.result_type,
                bool_label(result.detected),
                bool_label(result.succeeded),
                format_number(result.point_x),
                format_number(result.point_y),
                format_number(result.circle_x),
                format_number(result.circle_y),
                format_number(result.circle_r),
                format_number(result.circle_x_px),
                format_number(result.circle_y_px),
                format_number(result.circle_r_px),
                result.debug,
            ]
        )
    return rows


def bool_label(value: bool | None) -> str:
    if value is None:
        return ""
    return "yes" if value else "no"


def format_number(value: float | None) -> float | str:
    if value is None:
        return "-"
    return round(value, 6)


def percent(numerator: int, denominator: int) -> str:
    if denominator == 0:
        return "0.0%"
    return f"{numerator / denominator * 100:.1f}%"


def write_xlsx(path: Path, sheets: list[tuple[str, list[list[Any]]]]) -> None:
    names = unique_sheet_names([name for name, _rows in sheets])
    now = dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat()

    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        workbook.writestr("_rels/.rels", package_rels_xml())
        workbook.writestr("docProps/core.xml", core_props_xml(now))
        workbook.writestr("docProps/app.xml", app_props_xml(names))
        workbook.writestr("xl/workbook.xml", workbook_xml(names))
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        workbook.writestr("xl/styles.xml", styles_xml())
        for index, (_name, rows) in enumerate(sheets, start=1):
            workbook.writestr(f"xl/worksheets/sheet{index}.xml", worksheet_xml(rows))


def unique_sheet_names(names: list[str]) -> list[str]:
    used: set[str] = set()
    unique: list[str] = []
    for name in names:
        base = sanitize_sheet_name(name)
        candidate = base
        counter = 2
        while candidate.lower() in used:
            suffix = f" {counter}"
            candidate = f"{base[:31 - len(suffix)]}{suffix}"
            counter += 1
        used.add(candidate.lower())
        unique.append(candidate)
    return unique


def sanitize_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", " ", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return (cleaned or "Sheet")[:31]


def worksheet_xml(rows: list[list[Any]]) -> str:
    max_cols = max((len(row) for row in rows), default=1)
    max_rows = max(len(rows), 1)
    dimension = f"A1:{column_name(max_cols)}{max_rows}"
    xml: list[str] = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        f'<dimension ref="{dimension}"/>',
        '<sheetViews><sheetView workbookViewId="0">',
        '<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>',
        '<selection pane="bottomLeft"/>',
        "</sheetView></sheetViews>",
        columns_xml(rows),
        "<sheetData>",
    ]

    for row_index, row in enumerate(rows, start=1):
        xml.append(f'<row r="{row_index}">')
        for col_index, value in enumerate(row, start=1):
            is_table_header = bool(row) and row[0] in {"File name", "Model"}
            next_is_table_header = (
                row_index < len(rows)
                and bool(rows[row_index])
                and rows[row_index][0] in {"File name", "Model"}
            )
            is_section_title = sum(value not in (None, "") for value in row) == 1
            style = (
                2
                if is_table_header
                else 1
                if row_index == 1 or next_is_table_header or is_section_title
                else 0
            )
            xml.append(cell_xml(row_index, col_index, value, style))
        xml.append("</row>")

    xml.extend(["</sheetData>"])
    header_rows = [
        index
        for index, row in enumerate(rows, start=1)
        if row and row[0] in {"File name", "Model"}
    ]
    if len(header_rows) == 1 and max_rows > header_rows[0] and max_cols > 1:
        xml.append(
            f'<autoFilter ref="A{header_rows[0]}:{column_name(max_cols)}{max_rows}"/>'
        )
    xml.append("</worksheet>")
    return "".join(xml)


def columns_xml(rows: list[list[Any]]) -> str:
    widths: list[int] = []
    max_cols = max((len(row) for row in rows), default=0)
    for col_index in range(max_cols):
        max_width = 10
        for row in rows:
            if col_index < len(row):
                max_width = max(max_width, min(60, len(str(row[col_index])) + 2))
        widths.append(max_width)

    if not widths:
        return ""

    parts = ["<cols>"]
    for index, width in enumerate(widths, start=1):
        parts.append(f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>')
    parts.append("</cols>")
    return "".join(parts)


def cell_xml(row_index: int, col_index: int, value: Any, style: int) -> str:
    reference = f"{column_name(col_index)}{row_index}"
    style_attr = f' s="{style}"' if style else ""

    if value is None or value == "":
        return f'<c r="{reference}"{style_attr}/>'

    if isinstance(value, bool):
        return f'<c r="{reference}" t="b"{style_attr}><v>{1 if value else 0}</v></c>'

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{reference}"{style_attr}><v>{value}</v></c>'

    return (
        f'<c r="{reference}" t="inlineStr"{style_attr}>'
        f"<is><t>{escape(str(value))}</t></is></c>"
    )


def column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name or "A"


def content_types_xml(sheet_count: int) -> str:
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, sheet_count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f"{overrides}</Types>"
    )


def package_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def core_props_xml(timestamp: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>make_numbers_file.py</dc:creator>"
        "<cp:lastModifiedBy>make_numbers_file.py</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def app_props_xml(sheet_names: list[str]) -> str:
    titles = "".join(f"<vt:lpstr>{escape(name)}</vt:lpstr>" for name in sheet_names)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>Python</Application>"
        '<HeadingPairs><vt:vector size="2" baseType="variant">'
        "<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>"
        f"<vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant>"
        "</vt:vector></HeadingPairs>"
        f'<TitlesOfParts><vt:vector size="{len(sheet_names)}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>'
        "</Properties>"
    )


def workbook_xml(sheet_names: list[str]) -> str:
    sheets_xml = "".join(
        f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, name in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<workbookPr date1904=\"false\"/>"
        f"<sheets>{sheets_xml}</sheets>"
        "</workbook>"
    )


def workbook_rels_xml(sheet_count: int) -> str:
    rels = "".join(
        f'<Relationship Id="rId{index}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, sheet_count + 1)
    )
    styles_id = sheet_count + 1
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{rels}"
        f'<Relationship Id="rId{styles_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        "</Relationships>"
    )


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'
        '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'
        "</fonts>"
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/><bgColor indexed="64"/></patternFill></fill>'
        "</fills>"
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        "</cellXfs>"
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
        "</styleSheet>"
    )


def export_numbers(xlsx_path: Path, numbers_path: Path, force: bool) -> None:
    if sys.platform != "darwin":
        raise SystemExit("--numbers requires macOS with Apple Numbers installed")
    if shutil.which("osascript") is None:
        raise SystemExit("--numbers requires osascript")
    if numbers_path.exists():
        if not force:
            raise SystemExit(f"{numbers_path} already exists. Use --force to overwrite it.")
        if numbers_path.is_dir():
            shutil.rmtree(numbers_path)
        else:
            numbers_path.unlink()

    numbers_path.parent.mkdir(parents=True, exist_ok=True)
    script = f"""
set sourcePath to {applescript_string(str(xlsx_path))}
set outputPath to {applescript_string(str(numbers_path))}
tell application "Numbers"
    activate
    set importedDocument to open POSIX file sourcePath
    delay 1
    save importedDocument in POSIX file outputPath
    close importedDocument saving no
end tell
"""
    subprocess.run(["osascript", "-e", script], check=True)


def applescript_string(value: str) -> str:
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


if __name__ == "__main__":
    raise SystemExit(main())
