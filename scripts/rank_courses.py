#!/usr/bin/env python3
"""Deterministically rank courses from the MAcc exit survey workbook."""

from __future__ import annotations

import argparse
import csv
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
CELL_REF_RE = re.compile(r"([A-Z]+)")
CORE_PROMPT = "Please place each MAcc CORE course into rank order"
ELECTIVE_PROMPT = "Rate ACC"


@dataclass(frozen=True)
class CourseAggregate:
    course: str
    responses: int
    metric: float


def col_to_index(col: str) -> int:
    idx = 0
    for char in col:
        idx = idx * 26 + (ord(char) - ord("A") + 1)
    return idx - 1


def load_sheet_rows(xlsx_path: Path) -> list[dict[int, str]]:
    with zipfile.ZipFile(xlsx_path) as workbook_zip:
        shared_strings = []
        shared_xml = ET.fromstring(workbook_zip.read("xl/sharedStrings.xml"))
        for item in shared_xml.findall("a:si", NS):
            shared_strings.append("".join((text.text or "") for text in item.findall(".//a:t", NS)))

        sheet_xml = ET.fromstring(workbook_zip.read("xl/worksheets/sheet1.xml"))

    rows: list[dict[int, str]] = []
    for row in sheet_xml.findall(".//a:sheetData/a:row", NS):
        parsed_row: dict[int, str] = {}
        for cell in row.findall("a:c", NS):
            ref = cell.attrib.get("r", "")
            match = CELL_REF_RE.match(ref)
            if not match:
                continue
            col_idx = col_to_index(match.group(1))

            value_node = cell.find("a:v", NS)
            if value_node is None:
                parsed_row[col_idx] = ""
                continue

            raw = value_node.text or ""
            if cell.attrib.get("t") == "s":
                parsed_row[col_idx] = shared_strings[int(raw)]
            else:
                parsed_row[col_idx] = raw
        rows.append(parsed_row)
    return rows


def extract_course_name(header_text: str) -> str:
    return header_text.split(" - ")[-1].strip()


def aggregate_core_rankings(headers: dict[int, str], responses: list[dict[int, str]]) -> list[CourseAggregate]:
    course_cols = {
        col: extract_course_name(text)
        for col, text in headers.items()
        if CORE_PROMPT.lower() in text.lower()
    }

    totals = {course: 0.0 for course in course_cols.values()}
    counts = {course: 0 for course in course_cols.values()}

    for row in responses:
        for col, course in course_cols.items():
            value = row.get(col, "").strip()
            if value.isdigit():
                totals[course] += int(value)
                counts[course] += 1

    aggregate = [
        CourseAggregate(course=course, responses=counts[course], metric=totals[course] / counts[course])
        for course in course_cols.values()
        if counts[course] > 0
    ]
    return sorted(aggregate, key=lambda item: (item.metric, item.course))


def aggregate_elective_ratings(headers: dict[int, str], responses: list[dict[int, str]]) -> list[CourseAggregate]:
    course_cols = {
        col: extract_course_name(text)
        for col, text in headers.items()
        if ELECTIVE_PROMPT.lower() in text.lower()
    }

    totals = {course: 0.0 for course in course_cols.values()}
    counts = {course: 0 for course in course_cols.values()}

    for row in responses:
        for col, course in course_cols.items():
            value = row.get(col, "").strip()
            if value.isdigit():
                rating = int(value)
                if 1 <= rating <= 5:
                    totals[course] += rating
                    counts[course] += 1

    aggregate = [
        CourseAggregate(course=course, responses=counts[course], metric=totals[course] / counts[course])
        for course in course_cols.values()
        if counts[course] > 0
    ]
    return sorted(aggregate, key=lambda item: (-item.metric, item.course))


def write_csv(path: Path, rows: list[CourseAggregate], metric_name: str) -> None:
    with path.open("w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["rank", "course", metric_name, "responses"])
        for rank, item in enumerate(rows, start=1):
            writer.writerow([rank, item.course, f"{item.metric:.3f}", item.responses])


def write_markdown(path: Path, core_rows: list[CourseAggregate], elective_rows: list[CourseAggregate]) -> None:
    def as_table(rows: list[CourseAggregate], metric_name: str) -> str:
        lines = ["| Rank | Course | " + metric_name + " | Responses |", "|---:|---|---:|---:|"]
        for rank, item in enumerate(rows, start=1):
            lines.append(f"| {rank} | {item.course} | {item.metric:.3f} | {item.responses} |")
        return "\n".join(lines)

    report = [
        "# Exit Survey Course Ranking",
        "",
        "This report is generated deterministically from `Grad Program Exit Survey Data 2024.xlsx`.",
        "",
        "## Core course preference ranking",
        "Lower average rank is better (1 = most beneficial).",
        "",
        as_table(core_rows, "Average rank"),
        "",
        "## Elective course rating ranking",
        "Higher average rating is better (5 = most beneficial).",
        "",
        as_table(elective_rows, "Average rating"),
        "",
    ]
    path.write_text("\n".join(report), encoding="utf-8")


def write_summary(path: Path, core_rows: list[CourseAggregate], elective_rows: list[CourseAggregate]) -> None:
    payload = {
        "core_preference_ranking": [item.__dict__ for item in core_rows],
        "elective_rating_ranking": [item.__dict__ for item in elective_rows],
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", default="Grad Program Exit Survey Data 2024.xlsx")
    parser.add_argument("--output-dir", default="output")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    rows = load_sheet_rows(Path(args.input))
    if len(rows) < 4:
        raise ValueError("Workbook does not contain expected Qualtrics header rows and responses.")

    header_text_row = rows[1]
    response_rows = rows[3:]

    core_ranking = aggregate_core_rankings(header_text_row, response_rows)
    elective_ranking = aggregate_elective_ratings(header_text_row, response_rows)

    write_csv(output_dir / "core_course_ranking.csv", core_ranking, "average_rank")
    write_csv(output_dir / "elective_course_ranking.csv", elective_ranking, "average_rating")
    write_markdown(output_dir / "course_ranking_report.md", core_ranking, elective_ranking)
    write_summary(output_dir / "course_ranking_summary.json", core_ranking, elective_ranking)


if __name__ == "__main__":
    main()
