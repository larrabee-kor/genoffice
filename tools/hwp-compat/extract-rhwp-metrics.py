#!/usr/bin/env python3
"""Extract Hancom font advance widths from rhwp's metric tables into JSON.

rhwp (https://github.com/edwardkim/rhwp, MIT, Copyright (c) 2025-2026
Edward Kim) lays out HWP documents with per-font advance tables: generated
Latin/symbol ranges plus Hangul syllable groups, and an overlay region of
values measured from Hancom-exported PDFs. HWP/HWPX documents name those
fonts (휴먼명조, 한양신명조, HY헤드라인M…) and Hancom breaks their lines with
these widths, so the HWP compat faces (build-hwp-compat-fonts.py) carry them.

Reads rhwp's src/renderer/font_metrics_generated.rs and
src/renderer/font_metrics_overlays.rs (first matching entry wins, generated
before overlays, exactly like rhwp's FONT_METRICS view) and its
assets/font-rules/font_rule_registry_v2.json for Korean names that resolve
to a metric entry. Writes only the faces in FACES.

Usage: python3 tools/hwp-compat/extract-rhwp-metrics.py <rhwp checkout> [out.json]
"""

import json
import re
import sys
from pathlib import Path

DEFAULT_OUT = "tools/hwp-compat/hancom-metrics.json"

# HWP face name -> class of the outline source; names seen in public HWP/HWPX
# documents (Korean government forms and notices), resolved through rhwp
FACES = {
    "휴먼명조": "serif",
    "휴먼고딕": "sans",
    "한양신명조": "serif",
    "한양중고딕": "sans",
    "한양견명조": "serif",
    "한양견고딕": "sans",
    "신명조": "serif",
    "한컴바탕": "serif",
    "한컴돋움": "sans",
    "함초롬바탕": "serif",
    "함초롬돋움": "sans",
    "HY신명조": "serif",
    "HY중고딕": "sans",
    "HY견명조": "serif",
    "HY견고딕": "sans",
    "HY헤드라인M": "sans",
    "HY그래픽": "sans",
    "HY궁서": "serif",
    "한컴 윤고딕 720": "sans",
    "한컴 윤고딕 740": "sans",
    "한컴 윤고딕 760": "sans",
    "한컴 고딕": "sans",
    "신명 중명조": "serif",
    # Windows faces as Hancom lays them out (their own advances; Word's
    # substitution rules do not apply to an HWP-origin document)
    "바탕": "serif",
    "바탕체": "mono",
    "돋움": "sans",
    "돋움체": "mono",
    "굴림": "sans",
    "굴림체": "mono",
    "궁서": "serif",
    "궁서체": "mono",
    "새굴림": "sans",
    "맑은 고딕": "sans",
    # Hancom's Latin face, laid out with the face rhwp's registry names for it
    # (style-fallback): its widths feed line fitting, its glyphs come from
    # that installed face (build-hwp-compat-fonts.py LATIN_ALIASES)
    "HCI Poppy": "latin",
}

# Latin faces rhwp lays out with another face's table (registry style-fallback)
LATIN_LAYOUT = {"HCI Poppy": "Palatino Linotype"}

ARRAY_RE = re.compile(r"static (\w+): \[(u16|u8); \d+\] = \[([^\]]*)\];", re.S)
RANGES_RE = re.compile(r"static (\w+): \[LatinRange; \d+\] = \[(.*?)\];", re.S)
RANGE_RE = re.compile(r"LatinRange\s*\{\s*start:\s*(0x[0-9A-Fa-f]+),\s*end:\s*(0x[0-9A-Fa-f]+),\s*widths:\s*&(\w+),?\s*\}")
HANGUL_RE = re.compile(r"static (\w+): HangulMetric = HangulMetric \{(.*?)\};", re.S)
METRIC_RE = re.compile(
    r'FontMetric\s*\{\s*name:\s*"([^"]+)",\s*bold:\s*(true|false),\s*italic:\s*(true|false),\s*'
    r"em_size:\s*(\d+),\s*latin_ranges:\s*&(\w+),\s*hangul:\s*(None|Some\(&(\w+)\)),?\s*\}",
    re.S,
)


def parse(sources: list[str]):
    arrays, ranges, hanguls, metrics = {}, {}, {}, []
    for text in sources:
        for name, _, body in ARRAY_RE.findall(text):
            arrays[name] = [int(x) for x in re.findall(r"\d+", body)]
        for name, body in RANGES_RE.findall(text):
            ranges[name] = [(int(a, 16), int(b, 16), arr) for a, b, arr in RANGE_RE.findall(body)]
        for name, body in HANGUL_RE.findall(text):
            fields = dict(re.findall(r"(\w+):\s*&?(\w+)", body))
            hanguls[name] = fields
        for m in METRIC_RE.finditer(text):
            metrics.append(
                {
                    "name": m.group(1),
                    "bold": m.group(2) == "true",
                    "italic": m.group(3) == "true",
                    "em": int(m.group(4)),
                    "ranges": m.group(5),
                    "hangul": m.group(7),
                }
            )
    return arrays, ranges, hanguls, metrics


def resolve_aliases(registry: dict) -> dict[str, str]:
    alias = {}
    for rule in registry["rules"]:
        if rule.get("status") != "active":
            continue
        projections = {p["id"] for p in rule["projections"]}
        if "rust-layout-metric" in projections:
            alias.setdefault(rule["sourceFace"], rule["targetFaceOrPolicy"])
    return alias


def hangul_widths(h: dict, arrays: dict) -> list[int]:
    """distinct syllable widths (rhwp groups syllables by initial/medial/final)"""
    return sorted(set(arrays[h["widths"]]))


def face_table(metric: dict, arrays: dict, ranges: dict, hanguls: dict) -> dict:
    em = metric["em"]
    advances: dict[int, int] = {}
    for start, end, arr in ranges[metric["ranges"]]:
        widths = arrays[arr]
        for cp in range(start, end + 1):
            i = cp - start
            # first range wins; a zero width means "not covered" in rhwp
            if cp in advances or i >= len(widths) or widths[i] == 0:
                continue
            advances[cp] = widths[i]
    out = {
        "em": em,
        "advances": {f"{cp:04X}": round(w * 1000 / em) for cp, w in sorted(advances.items())},
    }
    if metric["hangul"]:
        ws = hangul_widths(hanguls[metric["hangul"]], arrays)
        out["hangul"] = round(sum(ws) / len(ws) * 1000 / em)
        if len(ws) > 1:
            out["hangulSpread"] = [round(min(ws) * 1000 / em), round(max(ws) * 1000 / em)]
    return out


def main() -> None:
    root = Path(sys.argv[1])
    out_path = Path(sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUT)
    renderer = root / "src" / "renderer"
    sources = [
        (renderer / "font_metrics_generated.rs").read_text(encoding="utf8"),
        (renderer / "font_metrics_overlays.rs").read_text(encoding="utf8"),
    ]
    arrays, ranges, hanguls, metrics = parse(sources)
    registry = json.loads(
        (root / "assets" / "font-rules" / "font_rule_registry_v2.json").read_text(encoding="utf8")
    )
    alias = resolve_aliases(registry)

    def find(name: str, bold: bool):
        for m in metrics:
            if m["name"] == name and m["bold"] == bold and not m["italic"]:
                return m
        return None

    faces = {}
    for face, cls in FACES.items():
        target = face if find(face, False) else (alias.get(face) or LATIN_LAYOUT.get(face))
        regular = find(target, False) if target else None
        if not regular:
            print(f"skip {face}: no rhwp metric", file=sys.stderr)
            continue
        entry = {"class": cls, "rhwpEntry": target, "regular": face_table(regular, arrays, ranges, hanguls)}
        bold = find(target, True)
        if bold:
            entry["bold"] = face_table(bold, arrays, ranges, hanguls)
        faces[face] = entry
    doc = {
        "source": "rhwp src/renderer/font_metrics_generated.rs + font_metrics_overlays.rs, "
        "assets/font-rules/font_rule_registry_v2.json (https://github.com/edwardkim/rhwp)",
        "license": "MIT, Copyright (c) 2025-2026 Edward Kim",
        "units": "advance widths in 1/1000 em; hangul = syllable advance",
        "faces": faces,
    }
    out_path.write_text(json.dumps(doc, ensure_ascii=False, indent=1) + "\n", encoding="utf8")
    print(f"{len(faces)} faces -> {out_path}")


if __name__ == "__main__":
    main()
