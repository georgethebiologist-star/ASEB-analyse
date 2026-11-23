"""Create Dutch SVG graphs from the master workbook.

Run `python src/waste_analysis.py`. Each event gets its own results folder;
`results/combined` contains totals across all events.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from html import escape
from math import cos, pi, sin
from pathlib import Path
import re

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "data" / "waste_scans_master.xlsx"
RESULTS = ROOT / "results"
STREAMS = ("GFT", "PD", "Paper", "Residual")
MATERIALS = ("GFT", "PD", "Papier", "Rest", "Wegwerp", "Metaal")
CORPORATE = ("#479456", "#E9BE32", "#7299CA", "#0A3A40", "#CC76A0", "#B8BCBA")
STREAM_COLORS = ("#479456", "#E9BE32", "#7299CA", "#CC76A0")
BACKGROUND, INK, MUTED = "#F0EEE8", "#0A3A40", "#586B6D"
CAPACITY_L = {"s": 30, "m": 70, "medium zwart": 70, "l": 140, "l+": 140,
              "large+": 140, "xl": 350, "xl= 350l": 350}
MONTHS = ("januari", "februari", "maart", "april", "mei", "juni", "juli",
          "augustus", "september", "oktober", "november", "december")
GROUPS = {"Employees": "medewerkers", "Students": "studenten", "Restaurant": "restaurant",
          "BG": "begane grond", "1e Verdieping": "eerste verdieping"}


@dataclass(frozen=True)
class Event:
    name: str
    start: date
    end: date


@dataclass(frozen=True)
class Measurement:
    event: str
    institution: str
    campus: str
    building: str
    environment: str
    stream: str
    total_kg: float
    parts_kg: tuple[float, ...]
    volume_l: float | None
    data_class: str = "A_measured"  # Local flag; Greenhub uses grade A + measured status separately.


def as_date(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])


def normalized_building(event: str, institution: str, campus: str, environment: str, building: str) -> str:
    # Source subtotals group these six student bag labels under SP904.
    if (event == "ASEB Feb 10-13 2026" and institution == "UvA" and campus == "Sciencepark"
            and environment == "Students" and building in {"SP906", "SP907", "SP909", "SP910", "SP912", "SP913"}):
        return "SP904"
    return building


def load_data() -> tuple[list[Event], list[Measurement]]:
    wb = openpyxl.load_workbook(WORKBOOK, read_only=True, data_only=True)
    events = [Event(str(r[0]), as_date(r[1]), as_date(r[2]))
              for r in wb["Events"].iter_rows(min_row=2, values_only=True) if r[0]]
    names = {e.name for e in events}
    if len(names) != len(events):
        raise ValueError("Event names must be unique")
    records: list[Measurement] = []
    for rownum, r in enumerate(wb["Scans"].iter_rows(min_row=2, values_only=True), 2):
        if not any(v is not None and str(v).strip() for v in r[:18]):
            continue
        event, institution, campus, building, environment, stream = (str(x or "").strip() for x in r[:6])
        if event not in names or not institution or not building or not environment or stream not in STREAMS:
            raise ValueError(f"Scans row {rownum}: complete the event, location, environment and stream")
        if any(x is None or x == "" for x in r[8:15]):
            raise ValueError(f"Scans row {rownum}: enter total and all six material weights in grams")
        total_g = float(r[8])
        parts_g = tuple(float(x) for x in r[9:15])
        if total_g <= 0 or any(x < 0 for x in parts_g) or abs(total_g - sum(parts_g)) > 20:
            raise ValueError(f"Scans row {rownum}: weights must be non-negative and close within 20 g")
        size = str(r[6] or "").strip().lower()
        fill = float(r[7]) if r[7] is not None else None
        if fill is not None and not 0 <= fill <= 1:
            raise ValueError(f"Scans row {rownum}: bag fill must be between 0 and 1")
        volume = CAPACITY_L[size] * fill if size in CAPACITY_L and fill is not None else None
        building = normalized_building(event, institution, campus, environment, building)
        records.append(Measurement(event, institution, campus, building, environment, stream,
                                   total_g / 1000, tuple(x / 1000 for x in parts_g), volume))
    if "Summary inputs" in wb.sheetnames:
        for rownum, r in enumerate(wb["Summary inputs"].iter_rows(min_row=2, values_only=True), 2):
            if not r[0]:
                continue
            event, institution, campus, building, environment, stream = (str(x or "").strip() for x in r[:6])
            if event not in names or stream not in STREAMS:
                raise ValueError(f"Summary inputs row {rownum}: invalid event or stream")
            total_g = float(r[7])
            parts_g = tuple(float(x or 0) for x in r[8:14])
            if total_g <= 0 or abs(total_g - sum(parts_g)) > 20:
                raise ValueError(f"Summary inputs row {rownum}: component weights do not match total")
            records.append(Measurement(event, institution, campus, building, environment, stream,
                                       total_g / 1000, tuple(x / 1000 for x in parts_g),
                                       float(r[6]) if r[6] is not None else None))
    return events, records


def date_label(events: list[Event]) -> str:
    def one(e: Event) -> str:
        if e.start.month == e.end.month and e.start.year == e.end.year:
            return f"{e.start.day}–{e.end.day} {MONTHS[e.start.month - 1]} {e.start.year}"
        return f"{e.start.day} {MONTHS[e.start.month - 1]} – {e.end.day} {MONTHS[e.end.month - 1]} {e.end.year}"
    return "  |  ".join(one(e) for e in events)


def slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def base(width: int, height: int, title: str, date_text: str) -> list[str]:
    return [f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">',
            f'<rect width="{width}" height="{height}" fill="{BACKGROUND}"/>',
            f'<text x="55" y="67" font-family="Arial" font-size="38" fill="{INK}">{escape(title)}</text>',
            f'<text x="57" y="110" font-family="Arial" font-size="23" fill="{MUTED}">{escape(date_text)}</text>']


def finish(folder: Path, filename: str, svg: list[str]) -> None:
    svg.append("</svg>")
    (folder / filename).write_text("\n".join(svg), encoding="utf-8")


def legend(svg: list[str], names: tuple[str, ...], colors: tuple[str, ...], y: int, width: int,
           line_name: str | None = None) -> None:
    items = list(zip(names, colors))
    gap = min(240, (width - 100) // max(1, len(items) + bool(line_name)))
    for i, (name, color) in enumerate(items):
        x = 55 + i * gap
        svg += [f'<rect x="{x}" y="{y-20}" width="21" height="21" fill="{color}"/>',
                f'<text x="{x+30}" y="{y-2}" font-family="Arial" font-size="22" fill="{INK}">{escape(name)}</text>']
    if line_name:
        x = 55 + len(items) * gap
        svg += [f'<line x1="{x}" y1="{y-10}" x2="{x+24}" y2="{y-10}" stroke="{INK}" stroke-width="5"/>',
                f'<text x="{x+32}" y="{y-2}" font-family="Arial" font-size="22" fill="{INK}">{escape(line_name)}</text>']


def horizontal_stacks(folder: Path, filename: str, title: str, dates: str,
                      rows: list[tuple[str, tuple[float, ...]]], names: tuple[str, ...],
                      colors: tuple[str, ...], unit: str = "kg") -> None:
    rows = [(label, values) for label, values in rows if sum(values) > 0]
    if not rows:
        return
    width, left, right, row_h = 1700, 350, 150, 58
    height = 205 + row_h * len(rows) + 105
    plot_w = width - left - right
    maximum = max(sum(v) for _, v in rows)
    suffix = "%" if unit == "%" else ""
    svg = base(width, height, title, dates)
    for tick in range(6):
        x = left + tick * plot_w / 5
        val = maximum * tick / 5
        svg += [f'<line x1="{x:.1f}" y1="165" x2="{x:.1f}" y2="{height-120}" stroke="#D1D4D0"/>',
                f'<text x="{x:.1f}" y="{height-83}" text-anchor="middle" font-family="Arial" font-size="19" fill="{MUTED}">{val:.1f}{suffix}</text>']
    for i, (label, values) in enumerate(rows):
        y = 185 + i * row_h
        svg.append(f'<text x="{left-18}" y="{y+28}" text-anchor="end" font-family="Arial" font-size="23" fill="{INK}">{escape(label)}</text>')
        x = left
        for value, color in zip(values, colors):
            bar_w = value / maximum * plot_w
            if bar_w > 0:
                svg.append(f'<rect x="{x:.2f}" y="{y}" width="{bar_w:.2f}" height="35" fill="{color}"/>')
            x += bar_w
        svg.append(f'<text x="{min(x+12,width-110):.1f}" y="{y+27}" font-family="Arial" font-size="20" fill="{INK}">{sum(values):.1f}{suffix}</text>')
    svg.append(f'<text x="{width-55}" y="{height-83}" text-anchor="end" font-family="Arial" font-size="20" fill="{MUTED}">{unit}</text>')
    legend(svg, names, colors, height-30, width)
    finish(folder, filename, svg)


def pie(folder: Path, filename: str, title: str, dates: str,
        names: tuple[str, ...], values: tuple[float, ...], colors: tuple[str, ...]) -> None:
    total = sum(values)
    if total <= 0:
        return
    width, height, cx, cy, radius = 1600, 880, 550, 490, 280
    svg = base(width, height, title, dates)
    angle = -pi / 2
    for value, color in zip(values, colors):
        turn = value / total * 2 * pi
        if value <= 0:
            continue
        if abs(value - total) < 1e-9:
            svg.append(f'<circle cx="{cx}" cy="{cy}" r="{radius}" fill="{color}"/>')
            break
        end = angle + turn
        x1, y1 = cx + radius * cos(angle), cy + radius * sin(angle)
        x2, y2 = cx + radius * cos(end), cy + radius * sin(end)
        large = 1 if turn > pi else 0
        svg.append(f'<path d="M {cx} {cy} L {x1:.2f} {y1:.2f} A {radius} {radius} 0 {large} 1 {x2:.2f} {y2:.2f} Z" fill="{color}" stroke="{BACKGROUND}" stroke-width="4"/>')
        angle = end
    for i, (name, value, color) in enumerate(zip(names, values, colors)):
        y = 260 + i * 82
        svg += [f'<rect x="1000" y="{y-24}" width="28" height="28" fill="{color}"/>',
                f'<text x="1044" y="{y}" font-family="Arial" font-size="27" fill="{INK}">{escape(name)}  {value/total*100:.1f}%</text>']
    finish(folder, filename, svg)


def gft_contamination(folder: Path, dates: str, group: str,
                      rows: list[tuple[str, tuple[float, ...]]],
                      total_by_location: dict[str, float]) -> None:
    rows = sorted(rows, key=lambda x: -x[1][0])
    if not rows or max(x[1][0] for x in rows) <= 0:
        return
    horizontal_stacks(folder, f"restafval_{slug(GROUPS.get(group, group))}_gft_kg.svg",
                      f"GFT in restafval: {GROUPS.get(group,group.lower())}", dates,
                      [(label, (parts[0],)) for label, parts in rows], ("GFT",), (CORPORATE[0],))
    percentages = sorted(
        ((label, (parts[0] / total_by_location[label] * 100,)) for label, parts in rows),
        key=lambda x: -x[1][0],
    )
    horizontal_stacks(folder, f"restafval_{slug(GROUPS.get(group, group))}_gft_procent.svg",
                      f"Aandeel GFT in restafval: {GROUPS.get(group,group.lower())}", dates,
                      percentages, ("GFT-aandeel",), (CORPORATE[0],), "%")


def location_rows(records: list[Measurement]) -> list[tuple[str, float, tuple[float, ...]]]:
    grouped = defaultdict(lambda: [0.0] * 7)
    for record in records:
        values = grouped[record.building]
        values[0] += record.total_kg
        for i, part in enumerate(record.parts_kg):
            values[i + 1] += part
    return [(name, values[0], tuple(values[1:])) for name, values in sorted(grouped.items())]


def matching_parts(stream: str, paper_in_rest: bool = False) -> tuple[int, ...]:
    if stream == "Residual":
        return (2, 3, 4, 5) if paper_in_rest else (3, 4, 5)
    return (STREAMS.index(stream),)


def paired_bars(folder: Path, filename: str, title: str, dates: str,
                rows: list[tuple[str, float, tuple[float, ...]]],
                stream: str, paper_in_rest: bool = False) -> None:
    """One total bar and one composition bar per location, with a purity line."""
    if not rows:
        return
    names = ("GFT", "PD", "Rest", "Wegwerp", "Metaal") if paper_in_rest else MATERIALS
    colors = (CORPORATE[0], CORPORATE[1], CORPORATE[3], CORPORATE[4], CORPORATE[5]) if paper_in_rest else CORPORATE
    target = matching_parts(stream, paper_in_rest)
    rows = sorted(rows, key=lambda row: -sum(row[2][i] for i in target) / row[1])
    width = max(1700, 320 + len(rows) * 190)
    height, top, bottom, left = 1000, 170, 650, 130
    plot_w = width - 260
    step = plot_w / len(rows)
    maximum = max(max(total, sum(parts)) for _, total, parts in rows) * 1.08
    svg = base(width, height, title, dates)
    for tick in range(6):
        y = bottom - tick * (bottom - top) / 5
        svg += [f'<line x1="{left}" y1="{y:.1f}" x2="{width-130}" y2="{y:.1f}" stroke="#D1D4D0"/>',
                f'<text x="{left-14}" y="{y+7:.1f}" text-anchor="end" font-family="Arial" font-size="20" fill="{MUTED}">{maximum*tick/5:.1f}</text>',
                f'<text x="{width-110}" y="{y+7:.1f}" font-family="Arial" font-size="20" fill="{MUTED}">{tick*20}%</text>']
    points = []
    for i, (name, total, parts) in enumerate(rows):
        center = left + step * (i + .5)
        bar_w = min(50, step * .24)
        gap = min(12, step * .05)
        total_h = total / maximum * (bottom - top)
        svg.append(f'<rect x="{center-gap/2-bar_w:.1f}" y="{bottom-total_h:.1f}" width="{bar_w:.1f}" height="{total_h:.1f}" fill="{STREAM_COLORS[STREAMS.index(stream)]}"/>')
        displayed = (parts[0], parts[1], parts[2]+parts[3], parts[4], parts[5]) if paper_in_rest else parts
        y = bottom
        for value, color in zip(displayed, colors):
            h = value / maximum * (bottom - top)
            if h > 0:
                y -= h
                svg.append(f'<rect x="{center+gap/2:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{h:.1f}" fill="{color}"/>')
        label_lines = [name] if len(name) <= 16 else [name[:16], name[16:]]
        for j, line in enumerate(label_lines):
            svg.append(f'<text x="{center:.1f}" y="{bottom+36+j*24}" text-anchor="middle" font-family="Arial" font-size="19" fill="{INK}">{escape(line)}</text>')
        share = sum(parts[j] for j in target) / total * 100
        points.append((center, bottom - share / 100 * (bottom - top)))
    svg.append('<polyline points="' + ' '.join(f'{x:.1f},{y:.1f}' for x,y in points) + f'" fill="none" stroke="{INK}" stroke-width="5"/>')
    for x, y in points:
        svg.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="6" fill="{INK}"/>')
    svg.append(f'<text x="{left}" y="{height-210}" font-family="Arial" font-size="20" fill="{MUTED}">kg</text>')
    legend(svg, names, colors, height-100, width)
    svg += [f'<rect x="55" y="{height-63}" width="21" height="21" fill="{STREAM_COLORS[STREAMS.index(stream)]}"/>',
            f'<text x="85" y="{height-45}" font-family="Arial" font-size="22" fill="{INK}">Totaal in bak</text>',
            f'<line x1="350" y1="{height-52}" x2="380" y2="{height-52}" stroke="{INK}" stroke-width="5"/>',
            f'<text x="390" y="{height-45}" font-family="Arial" font-size="22" fill="{INK}">{escape(("Restafval" if stream == "Residual" else ("Papier" if stream == "Paper" else stream)) + " in " + ("restafval" if stream == "Residual" else ("papier" if stream == "Paper" else stream)))}</text>']
    finish(folder, filename, svg)


def purity_extremes(folder: Path, prefix: str, dates: str, stream: str,
                    rows: list[tuple[str, float, tuple[float, ...]]], paper_in_rest: bool = False) -> None:
    if len(rows) < 2:
        return
    target = matching_parts(stream, paper_in_rest)
    ranked = sorted(rows, key=lambda row: sum(row[2][i] for i in target) / row[1])
    if abs(sum(ranked[-1][2][i] for i in target) / ranked[-1][1]
           - sum(ranked[0][2][i] for i in target) / ranked[0][1]) < 1e-9:
        return
    for word, (name, _, parts) in (("laagste", ranked[0]), ("hoogste", ranked[-1])):
        values = (parts[0], parts[1], parts[2]+parts[3], parts[4], parts[5]) if paper_in_rest else parts
        names = ("GFT", "PD", "Rest", "Wegwerp", "Metaal") if paper_in_rest else MATERIALS
        colors = (CORPORATE[0], CORPORATE[1], CORPORATE[3], CORPORATE[4], CORPORATE[5]) if paper_in_rest else CORPORATE
        pie(folder, f"{prefix}_{word}_zuiverheid.svg", f"{name}: {word} zuiverheid", dates,
            names, values, colors)


def mass_purity_scatter(folder: Path, filename: str, dates: str, stream: str,
                        rows: list[tuple[str, float, tuple[float, ...]]]) -> None:
    if len(rows) < 3:
        return
    target = matching_parts(stream)
    points = [(name, total, sum(parts[i] for i in target) / total * 100) for name, total, parts in rows]
    x_max = max(x for _, x, _ in points) * 1.1
    width, height, left, right, top, bottom = 1600, 900, 190, 150, 170, 700
    svg = base(width, height, f"Gewicht en zuiverheid: {('Restafval' if stream == 'Residual' else 'Papier' if stream == 'Paper' else stream)}", dates)
    for tick in range(6):
        x = left + (width-left-right) * tick / 5
        y = bottom - (bottom-top) * tick / 5
        svg += [f'<line x1="{x:.1f}" y1="{top}" x2="{x:.1f}" y2="{bottom}" stroke="#D1D4D0"/>',
                f'<line x1="{left}" y1="{y:.1f}" x2="{width-right}" y2="{y:.1f}" stroke="#D1D4D0"/>',
                f'<text x="{x:.1f}" y="{bottom+37}" text-anchor="middle" font-family="Arial" font-size="20" fill="{MUTED}">{x_max*tick/5:.1f}</text>',
                f'<text x="{left-20}" y="{y+6:.1f}" text-anchor="end" font-family="Arial" font-size="20" fill="{MUTED}">{tick*20}%</text>']
    mean_x = sum(x for _, x, _ in points) / len(points)
    mean_y = sum(y for _, _, y in points) / len(points)
    spread = sum((x - mean_x) ** 2 for _, x, _ in points)
    if spread > 0:
        slope = sum((x - mean_x) * (y - mean_y) for _, x, y in points) / spread
        intercept = mean_y - slope * mean_x
        x1, x2 = min(x for _, x, _ in points), max(x for _, x, _ in points)
        y1, y2 = intercept + slope * x1, intercept + slope * x2
        svg.append(f'<line x1="{left+x1/x_max*(width-left-right):.1f}" y1="{bottom-y1/100*(bottom-top):.1f}" x2="{left+x2/x_max*(width-left-right):.1f}" y2="{bottom-y2/100*(bottom-top):.1f}" stroke="{INK}" stroke-width="4"/>')
    for name, mass, share in points:
        x = left + mass / x_max * (width-left-right)
        y = bottom - share / 100 * (bottom-top)
        svg += [f'<circle cx="{x:.1f}" cy="{y:.1f}" r="10" fill="{STREAM_COLORS[STREAMS.index(stream)]}"/>',
                f'<text x="{x+14:.1f}" y="{y-12:.1f}" font-family="Arial" font-size="18" fill="{INK}">{escape(name)}</text>']
    svg.append(f'<text x="{width-right}" y="{height-55}" text-anchor="end" font-family="Arial" font-size="22" fill="{INK}">Gewicht in bak (kg)</text>')
    finish(folder, filename, svg)


def render_scope(folder: Path, dates: str, records: list[Measurement]) -> int:
    folder.mkdir(parents=True, exist_ok=True)
    for old in folder.glob("*.svg"):
        old.unlink()
    present = [stream for stream in ("GFT", "Paper", "PD", "Residual")
               if any(r.stream == stream for r in records)]
    environments = [name for name in ("Employees", "Students", "Restaurant") if any(r.environment == name for r in records)]
    environments += sorted({r.environment for r in records} - set(environments))
    def stream_totals(filename: str, title: str, use_volume: bool) -> None:
        data = defaultdict(lambda: [0.0] * 4)
        for r in records:
            value = r.volume_l if use_volume else r.total_kg
            if value is not None:
                data[r.building][STREAMS.index(r.stream)] += value
        rows = [(label, tuple(values)) for label, values in sorted(data.items())]
        if rows:
            horizontal_stacks(folder, filename, title, dates, rows,
                              ("GFT", "PD", "Papier", "Restafval"), STREAM_COLORS,
                              "L" if use_volume else "kg")

    # Overview: volume, weight, then contamination by material.
    stream_totals("totaal_volume_per_locatie.svg", "Totaal volume per locatie en afvalstroom", True)
    stream_totals("totaal_gewicht_per_locatie.svg", "Totaal gewicht per locatie en afvalstroom", False)
    contamination = []
    for stream in present:
        subset = [r for r in records if r.stream == stream]
        total = sum(r.total_kg for r in subset)
        accepted = matching_parts(stream)
        parts = tuple(sum(r.parts_kg[i] for r in subset) / total * 100 if i not in accepted else 0 for i in range(6))
        contamination.append((("Restafval" if stream == "Residual" else "Papier" if stream == "Paper" else stream), parts))
    horizontal_stacks(folder, "totaal_vervuiling_per_categorie.svg",
                      "Vervuiling per afvalstroom en categorie (%)", dates,
                      contamination, MATERIALS, CORPORATE, "%")

    # Tissue analysis is omitted: the workbook has no tissue stream or tissue material column.
    gft_elsewhere = [r for r in records if r.stream != "GFT" and r.parts_kg[0] > 0]
    if gft_elsewhere:
        sources = [stream for stream in STREAMS if stream != "GFT"
                   and any(r.stream == stream for r in gft_elsewhere)]
        values = tuple(sum(r.parts_kg[0] for r in gft_elsewhere if r.stream == stream) for stream in sources)
        names = tuple("Restafval" if s == "Residual" else "Papier" if s == "Paper" else s for s in sources)
        colors = tuple(STREAM_COLORS[STREAMS.index(s)] for s in sources)
        pie(folder, "gft_vervuiling_in_andere_stromen.svg", "GFT in andere afvalstromen", dates,
            names, values, colors)

    # Each available stream follows the same environment and location workflow.
    for stream in present:
        stream_records = [r for r in records if r.stream == stream]
        label = "Restafval" if stream == "Residual" else "Papier" if stream == "Paper" else stream
        prefix = slug(label)
        for environment in environments:
            subset = [r for r in stream_records if r.environment == environment]
            if not subset:
                continue
            env_label = GROUPS.get(environment, environment.lower())
            base_name = f"{prefix}_{slug(env_label)}"
            rows = location_rows(subset)
            paired_bars(folder, f"{base_name}_gewicht_en_samenstelling.svg",
                        f"{label}: {env_label} per locatie", dates, rows, stream)
            purity_extremes(folder, base_name, dates, stream, rows)
            if stream == "Residual":
                paired_bars(folder, f"{base_name}_papier_bij_rest.svg",
                            f"{label}: {env_label}, papier bij rest", dates, rows, stream, True)
                gft_rows = [(name, parts) for name, _, parts in rows]
                totals = {name: total for name, total, _ in rows}
                gft_contamination(folder, dates, environment, gft_rows, totals)
        if len(location_rows(stream_records)) >= 3:
            mass_purity_scatter(folder, f"{prefix}_gewicht_en_zuiverheid.svg", dates,
                                stream, location_rows(stream_records))
    return len(list(folder.glob("*.svg")))


def main() -> None:
    events, records = load_data()
    if not records:
        raise ValueError("No measured records found in the master workbook")
    folders = []
    for event in events:
        subset = [r for r in records if r.event == event.name]
        if subset:
            folder = RESULTS / slug(event.name)
            folders.append((folder, render_scope(folder,date_label([event]),subset)))
    folders.append((RESULTS / "combined", render_scope(RESULTS / "combined",date_label(events),records)))
    print("Created " + ", ".join(f"{n} graphs in {folder}" for folder,n in folders))


if __name__ == "__main__":
    main()
