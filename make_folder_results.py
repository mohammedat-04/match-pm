#!/usr/bin/env python3
"""Build the old/new, vacuum/corner workbook with model summaries."""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import json
import math
import re

import make_numbers_file as workbook


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "old_new_models_results.xlsx"
SCALE = 971.43 / 221.0
VACUUM_REFERENCE = (1294.0, 1038.0, 221.0)
CORNER_REFERENCE = (1327.25, 902.02)
TOLERANCE = 6.0
DATE_SUFFIX = re.compile(r"_\d{8}_\d{2}_\d{2}_\d{2}$")


def yes_no(value: bool) -> str:
    return "Yes" if value else "No"


def number(value: float | None, digits: int = 3) -> float | str:
    return "-" if value is None else round(value, digits)


def model_label(result: workbook.ParsedResult) -> str:
    if result.provider or result.model:
        return " / ".join(value for value in (result.provider, result.model) if value)

    name = DATE_SUFFIX.sub("", result.run_name)
    for prefix in ("vacuum_gripper_", "sensor_corner_"):
        if name.startswith(prefix):
            name = name[len(prefix) :]
            break
    provider, separator, model = name.partition("_")
    return f"{provider} / {model}" if separator else provider


def debug_reason(data: dict) -> str:
    messages = [str(item).strip() for item in data.get("process_debug", [])]
    return "; ".join(message for message in messages if message)


def experiment_group(result: workbook.ParsedResult) -> tuple[str, str]:
    """Return the normalized prompt version and specification type from the path."""
    if len(result.path.parts) < 3:
        return "Original Prompt", "Unknown"

    experiment_folder = (
        result.path.parts[2].lower().replace("-", "_").replace(" ", "_")
    )
    if result.path.parts[1] == "corner":
        prompt = "Original Prompt"
        specification_folder = (
            result.path.parts[3].lower()
            if experiment_folder == "old_models" and len(result.path.parts) > 3
            else experiment_folder
        )
        specification = "Name Only" if "name" in specification_folder else "Full Specs"
        return prompt, specification

    normalized_prompt = experiment_folder.replace(" ", "_")
    if normalized_prompt in {"new_prompt_1", "new_prompt1"}:
        prompt = "New Prompt 1"
    elif normalized_prompt == "new_prompt":
        prompt = "New Prompt"
    elif normalized_prompt in {"or_prompt", "original_prompt"}:
        prompt = "Original Prompt"
    else:
        prompt = result.path.parts[2].replace("_", " ").title()

    detail_folder = result.path.parts[3].lower() if len(result.path.parts) > 3 else ""
    specification = "Name Only" if "name" in detail_folder else "Full Specs"
    return prompt, specification


def vacuum_record(result: workbook.ParsedResult, data: dict) -> dict:
    circles = workbook.parse_circles(str(data.get("vision_result", "")))
    candidates = []
    for circle in circles:
        x_px = VACUUM_REFERENCE[0] + (8.79 - circle["x"]) / SCALE
        y_px = VACUUM_REFERENCE[1] + (circle["y"] - 290.11) / SCALE
        r_px = circle["r"] / SCALE
        succeeded = all(
            abs(value - reference) <= TOLERANCE
            for value, reference in zip((x_px, y_px, r_px), VACUUM_REFERENCE)
        )
        candidates.append((circle, x_px, y_px, r_px, succeeded))

    succeeded = any(candidate[4] for candidate in candidates)
    selected = next(
        (candidate for candidate in candidates if candidate[4]),
        candidates[0] if candidates else None,
    )
    if succeeded:
        failure_cause = ""
    elif not selected:
        if data.get("_missing_json"):
            failure_cause = "Missing vision_results.json"
        else:
            failure_cause = debug_reason(data) or (
                "vision_ok is false; no circle detected"
                if result.vision_ok is False
                else "No circle detected"
            )
    else:
        deltas = [
            abs(selected[index] - VACUUM_REFERENCE[index - 1])
            for index in range(1, 4)
        ]
        failed = [
            label
            for label, delta in zip(("x_px", "y_px", "r_px"), deltas)
            if delta > TOLERANCE
        ]
        failure_cause = (
            f"Outside ±{TOLERANCE:g} px tolerance: {', '.join(failed)}"
        )

    if selected:
        circle, x_px, y_px, r_px, _ = selected
        x, y, radius = circle["x"], circle["y"], circle["r"]
    else:
        x = y = radius = x_px = y_px = r_px = None

    return {
        "row": [
            str(result.path.parent),
            yes_no(result.vision_ok is True),
            number(x),
            number(y),
            number(radius),
            number(x_px),
            number(y_px),
            number(r_px),
            yes_no(bool(candidates)),
            yes_no(succeeded),
            yes_no(len(candidates) > 1),
            failure_cause,
        ],
        "detected": bool(candidates),
        "succeeded": succeeded,
        "many_results": len(candidates) > 1,
    }


def corner_record(result: workbook.ParsedResult, data: dict) -> dict:
    x, y = result.point_x, result.point_y
    detected = x is not None and y is not None
    if detected:
        x_px = 1296.0 - x / SCALE
        y_px = 972.0 + y / SCALE
        dx_px = x_px - CORNER_REFERENCE[0]
        dy_px = y_px - CORNER_REFERENCE[1]
        error_px = math.hypot(dx_px, dy_px)
        succeeded = error_px <= TOLERANCE
        failure_cause = (
            ""
            if succeeded
            else f"Pixel error {error_px:.3f} exceeds {TOLERANCE:g} px tolerance"
        )
    else:
        x_px = y_px = dx_px = dy_px = error_px = None
        succeeded = False
        if data.get("_missing_json"):
            failure_cause = "Missing vision_results.json"
        else:
            failure_cause = debug_reason(data) or (
                "vision_ok is false; no point detected"
                if result.vision_ok is False
                else "No point detected"
            )

    quality = data.get("quality_score", {}).get("cornerDetectionSubPixel", {}) or {}
    return {
        "row": [
            str(result.path.parent),
            yes_no(result.vision_ok is True),
            number(x),
            number(y),
            number(x_px),
            number(y_px),
            number(dx_px),
            number(dy_px),
            number(error_px),
            yes_no(succeeded),
            quality.get("residual_score", "-"),
            quality.get("inlier_score", "-"),
            quality.get("weakest_score_action", "-"),
            failure_cause,
        ],
        "detected": detected,
        "succeeded": succeeded,
        "many_results": False,
    }


def build_sheet(results, source_data, group: str, target: str):
    grouped = defaultdict(list)
    for result in results:
        logical_group = (
            "old models"
            if len(result.path.parts) > 2
            and result.path.parts[:3] == ("new models", "corner", "old models")
            else result.group
        )
        if (
            logical_group == group
            and len(result.path.parts) > 1
            and result.path.parts[1] == target
        ):
            prompt, specification = experiment_group(result)
            grouped[(prompt, specification, model_label(result))].append(result)

    is_vacuum = target == "vacuum_gripper"
    detail_header = (
        [
            "File name",
            "vision_ok",
            "x",
            "y",
            "r",
            "x_px",
            "y_px",
            "r_px",
            "Detected",
            "Succeeded",
            "Many results",
            "Failure cause",
        ]
        if is_vacuum
        else [
            "File name",
            "vision_ok",
            "x",
            "y",
            "x_px",
            "y_px",
            "dx_px",
            "dy_px",
            "error_px",
            "Succeeded",
            "residual_score",
            "inlier_score",
            "action",
            "Failure cause",
        ]
    )
    make_record = vacuum_record if is_vacuum else corner_record
    records_by_group = {}
    for experiment, model_results in grouped.items():
        records_by_group[experiment] = [
            make_record(result, source_data[result.path])
            for result in sorted(model_results, key=lambda item: str(item.path))
        ]

    prompt_order = {"Original Prompt": 0, "New Prompt": 1, "New Prompt 1": 2}
    specification_order = {"Name Only": 0, "Full Specs": 1}
    group_order = sorted(
        records_by_group,
        key=lambda key: (
            prompt_order.get(key[0], 99),
            specification_order.get(key[1], 99),
            key[2].casefold(),
        ),
    )
    summary_groups = []
    for prompt, specification, _model in group_order:
        pair = (prompt, specification)
        if pair not in summary_groups:
            summary_groups.append(pair)

    rows = []
    for prompt, specification in summary_groups:
        if rows:
            rows.append([])
        rows.extend(
            [
                [f"Results Summary — {prompt} — {specification}"],
                [
                    "Model",
                    "Total",
                    "Detected",
                    "Succeeded",
                    "Success rate",
                    "Many-result runs",
                ],
            ]
        )
        for experiment in group_order:
            experiment_prompt, experiment_specification, model = experiment
            if (experiment_prompt, experiment_specification) != (prompt, specification):
                continue
            records = records_by_group[experiment]
            total = len(records)
            succeeded = sum(record["succeeded"] for record in records)
            rows.append(
                [
                    model,
                    total,
                    sum(record["detected"] for record in records),
                    succeeded,
                    workbook.percent(succeeded, total),
                    sum(record["many_results"] for record in records),
                ]
            )

    rows.extend([[], []])
    for prompt, specification in summary_groups:
        rows.append([f"Detailed Results — {prompt} — {specification}"])
        for experiment in group_order:
            experiment_prompt, experiment_specification, model = experiment
            if (experiment_prompt, experiment_specification) != (prompt, specification):
                continue
            rows.extend([[model], detail_header])
            rows.extend(record["row"] for record in records_by_group[experiment])
            rows.append([])
        rows.append([])

    if not records_by_group:
        rows.extend([["No results found"], detail_header])
    return rows


def collect_run_results(calibration: workbook.Calibration):
    """Collect JSON results and add failed rows for run folders without JSON."""
    results = workbook.collect_results(ROOT, calibration)
    known_run_folders = {result.path.parent for result in results}

    run_patterns = [
        ("old models", "vacuum_gripper", "*/*/*"),
        ("old models", "corner", "*/*"),
        ("new models", "vacuum_gripper", "*/*/*"),
        ("new models", "corner", "Corner_*/*"),
        ("new models", "corner", "old models/*/*"),
    ]
    for group, target, pattern in run_patterns:
        target_folder = ROOT / group / target
        if not target_folder.is_dir():
            continue
        for run_folder in sorted(target_folder.glob(pattern)):
            if not run_folder.is_dir():
                continue
            relative_folder = run_folder.relative_to(ROOT)
            if relative_folder in known_run_folders:
                continue
            metadata = workbook.parse_run_metadata(run_folder.name)
            missing = workbook.missing_result(group, run_folder.name, metadata)
            missing.path = relative_folder / "vision_results.json"
            results.append(missing)
            known_run_folders.add(relative_folder)

    return sorted(results, key=lambda result: str(result.path))


def main() -> None:
    calibration = workbook.Calibration(
        SCALE,
        1296.0,
        972.0,
        *VACUUM_REFERENCE,
        TOLERANCE,
        TOLERANCE,
    )
    results = collect_run_results(calibration)
    source_data = {
        result.path: (
            json.loads((ROOT / result.path).read_text(encoding="utf-8"))
            if (ROOT / result.path).is_file()
            else {"_missing_json": True}
        )
        for result in results
    }
    sheets = [
        ("New - Vacuum", build_sheet(results, source_data, "new models", "vacuum_gripper")),
        ("New - Corner", build_sheet(results, source_data, "new models", "corner")),
        ("Old - Vacuum", build_sheet(results, source_data, "old models", "vacuum_gripper")),
        ("Old - Corner", build_sheet(results, source_data, "old models", "corner")),
    ]
    workbook.write_xlsx(OUTPUT, sheets)
    print(f"Wrote {OUTPUT} ({len(results)} results, {len(sheets)} sheets)")


if __name__ == "__main__":
    main()
