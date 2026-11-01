#!/usr/bin/env python3
"""Build the HWP compat faces: Noto outlines on Hancom's advance widths.

HWP/HWPX documents name Hancom faces (휴먼명조, 한양신명조, HY헤드라인M…) that are
never installed outside Hancom Office, and Hangul breaks their lines with
those faces' own advances: punctuation, digits and Latin differ per face
(휴먼명조 space 0.50 em and “ 1.00 em, 함초롬바탕 hangul 0.97 em). A converted
document drawn with the Word substitute chain therefore wraps one or two
syllables away from the original on many lines.

For every face in tools/hwp-compat/hancom-metrics.json (rhwp's tables, MIT,
see extract-rhwp-metrics.py) this writes one small woff2 holding the
non-Hangul glyphs of the face's coverage, outlines taken from the bundled
GenOffice Serif KR / Sans KR subsets (Noto CJK KR, OFL) and placed on the
Hancom advance (centred; opening marks keep their right side bearing,
closing marks their left one, as in full-width designs). Layout tables are
dropped so no locale feature swaps a glyph for an unadjusted one. Hangul
syllables reuse the 1 em KR subset through a second @font-face, scaled when
the face's syllables are narrower.

Outputs:
  apps/docs/src/renderer/fonts/hwp-compat/<slug>.woff2
  apps/docs/src/renderer/fonts/hwp-compat.css
  packages/hwpx2ooxml/src/compat-faces.ts   (face name -> compat family)

Usage: python3 tools/hwp-compat/build-hwp-compat-fonts.py
"""

import json
import re
import subprocess
from pathlib import Path

from fontTools.misc.transform import Transform
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.t2CharStringPen import T2CharStringPen
from fontTools.pens.transformPen import TransformPen
from fontTools.subset import Options, Subsetter
from fontTools.ttLib import TTFont

ROOT = Path(__file__).resolve().parent.parent.parent
FONTS = ROOT / "apps/docs/src/renderer/fonts"
METRICS = ROOT / "tools/hwp-compat/hancom-metrics.json"
OUT_DIR = FONTS / "hwp-compat"
OUT_CSS = FONTS / "hwp-compat.css"
OUT_TS = ROOT / "packages/hwpx2ooxml/src/compat-faces.ts"
OUT_METRICS = ROOT / "packages/hwpx2ooxml/src/compat-metrics.ts"

SOURCES = {
    "serif": "GenOfficeSerifKR-Regular-subset.woff2",
    "sans": "GenOfficeSansKR-Regular-subset.woff2",
    "mono": "GenOfficeSansKR-Regular-subset.woff2",
}

# ASCII names for faces whose rhwp entry is itself Korean
SLUGS = {
    "휴먼명조": "HumanMyeongjo",
    # the measured overlay entry; the face keeps its first published name
    "HumanMyeongJo": "HumanMyeongjo",
    "휴먼고딕": "HumanGothic",
    "한컴 윤고딕 720": "HancomYoonGothic720",
    "한컴 윤고딕 740": "HancomYoonGothic740",
    "한컴 윤고딕 760": "HancomYoonGothic760",
    "한컴 고딕": "HancomGothic",
    "신명 중명조": "SinmyeongJungmyeongjo",
}

# Marks Hangul draws from its own symbol face when the document face has no
# glyph: U+2027 (개정‧시행) measures 0.173 em in the corpus PDFs (median of 13
# isolated runs), as in rhwp's 한컴바탕 table (0.174); sans faces borrow 한컴돋움
BORROWED = {0x2027: {"serif": "한컴바탕", "sans": "한컴돋움"}}

# Per-face corrections to rhwp's tables. 휴먼명조's middle dot has two Hancom
# realizations (rhwp font_metrics_overlays #7092): 1.000 em where Hangul draws
# the TrueType face (n=36), 0.384 where it substitutes an HFT subset; rhwp
# keeps 0.30. The government PDFs in the corpus are TrueType ones (k289,
# k042 wrap like the original only at 1.000)
OVERRIDES = {"휴먼명조": {0x00B7: 1000}}

# Hangul sets a space half an em wide unless the character shape asks for the
# face's own space (charPr useFontSpace="1"; 2 of 3,015 shapes in the corpus):
# lawgo-23 indents with HY신명조 spaces of 0.48-0.50 em although the face's own
# space is 0.333 em. The en space carries the converter's head gaps and fixed-
# width spaces (hwpx2ooxml fixedSpaceRun), so it is half an em in every face.
HALF_EM_SPACE = {0x0020: 500, 0x00A0: 500, 0x2002: 500}

# Hancom's Latin faces lay out with the Latin face rhwp's registry names for
# them (font_rule_registry_v2 style-fallback): installed under that name on
# Windows (Palatino Linotype) and macOS (Palatino, Latin widths within 0.2%)
LATIN_ALIASES = {"HCI Poppy": ["Palatino Linotype", "Palatino"]}

OPENING = set("‘“([{〈《「『【〔")
CLOSING = set("’”)]}〉》」』】〕,.、。")
HANGUL_RANGE = "U+AC00-D7A3"


def slug_of(entry: str) -> str:
    return SLUGS.get(entry) or re.sub(r"[^A-Za-z0-9]", "", entry)


def unicode_ranges(cps: list[int]) -> str:
    out, start, prev = [], None, None
    for cp in sorted(cps):
        if start is None:
            start = prev = cp
        elif cp == prev + 1:
            prev = cp
        else:
            out.append((start, prev))
            start = prev = cp
    if start is not None:
        out.append((start, prev))
    return ", ".join(f"U+{a:04X}" if a == b else f"U+{a:04X}-{b:04X}" for a, b in out)


def build_face(slug: str, cls: str, table: dict) -> list[int]:
    font = TTFont(str(FONTS / SOURCES[cls]))
    upm = font["head"].unitsPerEm
    source_cmap = font.getBestCmap()
    wanted = {
        int(cp, 16): adv
        for cp, adv in table["advances"].items()
        if not 0xAC00 <= int(cp, 16) <= 0xD7A3 and int(cp, 16) in source_cmap
    }
    options = Options(notdef_outline=True, glyph_names=False, layout_features=[])
    options.drop_tables += ["GSUB", "GPOS", "GDEF", "BASE", "kern", "vhea", "vmtx", "VORG"]
    subsetter = Subsetter(options=options)
    subsetter.populate(unicodes=sorted(wanted))
    subsetter.subset(font)

    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    cff = font["CFF "].cff
    charstrings = cff[cff.fontNames[0]].CharStrings
    glyph_set = font.getGlyphSet()
    done = set()
    for cp, adv_em in wanted.items():
        name = cmap.get(cp)
        if name is None or name in done:
            continue
        done.add(name)
        adv = round(adv_em * upm / 1000)
        old_adv, old_lsb = hmtx[name]
        bp = BoundsPen(glyph_set)
        glyph_set[name].draw(bp)
        if bp.bounds is None:  # spaces
            hmtx[name] = (adv, 0)
            continue
        x_min, _, x_max, _ = bp.bounds
        ink = x_max - x_min
        char = chr(cp)
        if char in OPENING and adv > old_adv:
            target_min = adv - (old_adv - x_max) - ink
        elif char in CLOSING and adv > old_adv:
            target_min = x_min
        else:
            target_min = x_min + (adv - old_adv) / 2
        dx = target_min - x_min
        if dx:
            pen = T2CharStringPen(adv, glyph_set)
            glyph_set[name].draw(TransformPen(pen, Transform().translate(dx, 0)))
            charstrings[name] = pen.getCharString(
                private=charstrings[name].private, globalSubrs=cff.GlobalSubrs
            )
        hmtx[name] = (adv, round(target_min))

    family = f"HWP Compat {slug}"
    ps_name = f"HWPCompat-{slug}"
    for rec in list(font["name"].names):
        if rec.nameID in (1, 4, 16):
            font["name"].setName(family, rec.nameID, rec.platformID, rec.platEncID, rec.langID)
        elif rec.nameID in (3, 6):
            font["name"].setName(ps_name, rec.nameID, rec.platformID, rec.platEncID, rec.langID)
    # reproducible output: the source's timestamps, not the build time
    font.recalcTimestamp = False
    font.flavor = "woff2"
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    font.save(str(OUT_DIR / f"{slug}.woff2"))
    return sorted(cp for cp in wanted if cmap.get(cp))


def main() -> None:
    data = json.loads(METRICS.read_text(encoding="utf8"))
    built: dict[str, dict] = {}
    names: dict[str, str] = {}
    for face, entry in data["faces"].items():
        table = entry["regular"]
        if not table["advances"]:
            continue
        if entry["class"] == "latin":
            # drawn by an installed face (LATIN_ALIASES); only its widths are kept
            for cp, adv in HALF_EM_SPACE.items():
                table["advances"][f"{cp:04X}"] = adv
            continue
        for cp, adv in HALF_EM_SPACE.items():
            table["advances"][f"{cp:04X}"] = adv
        for cp, adv in OVERRIDES.get(face, {}).items():
            table["advances"][f"{cp:04X}"] = adv
        for cp, donors in BORROWED.items():
            donor = donors.get(entry["class"])
            key = f"{cp:04X}"
            if donor and key not in table["advances"]:
                table["advances"][key] = data["faces"][donor]["regular"]["advances"][key]
        slug = slug_of(entry["rhwpEntry"])
        if slug not in built:
            cps = build_face(slug, entry["class"], table)
            built[slug] = {"class": entry["class"], "cps": cps, "hangul": table.get("hangul", 1000)}
        names[face] = [f"HWP Compat {slug}"]
    names.update(LATIN_ALIASES)

    css = [
        "/* Generated by tools/hwp-compat/build-hwp-compat-fonts.py — do not edit.",
        " * HWP compat faces: Noto CJK KR outlines (OFL, via the GenOffice KR subsets)",
        " * on Hancom advance widths from rhwp's measured tables (MIT, see",
        " * tools/hwp-compat/hancom-metrics.json). Used only for documents converted",
        " * from HWP/HWPX (line-metrics.ts hwpCompatFontFamily). */",
    ]
    for slug, info in sorted(built.items()):
        family = f"HWP Compat {slug}"
        css.append(
            "@font-face {\n"
            f"  font-family: '{family}';\n"
            f"  src: url('./hwp-compat/{slug}.woff2') format('woff2');\n"
            f"  unicode-range: {unicode_ranges(info['cps'])};\n"
            "  font-display: block;\n"
            "}"
        )
        hangul = info["hangul"]
        adjust = f"\n  /* rhwp table: syllables {hangul / 1000:.3f} em */\n  size-adjust: {hangul / 10:g}%;" if hangul != 1000 else ""
        css.append(
            "@font-face {\n"
            f"  font-family: '{family}';\n"
            f"  src: url('./{SOURCES[info['class']]}') format('woff2');\n"
            f"  unicode-range: {HANGUL_RANGE};{adjust}\n"
            "  font-display: block;\n"
            "}"
        )
    OUT_CSS.write_text("\n".join(css) + "\n", encoding="utf8")

    ts = [
        "// Generated by tools/hwp-compat/build-hwp-compat-fonts.py — do not edit.",
        "/** HWP face name -> the faces carrying Hancom's advances (apps/docs/src/renderer/fonts/hwp-compat.css) */",
        "export const HWP_COMPAT_FACES: Readonly<Record<string, readonly string[]>> = {",
        *[f"  '{face}': [{', '.join(repr(f) for f in families)}]," for face, families in sorted(names.items())],
        "}",
        "",
        "/** the generic family closing a compat face's chain */",
        "export const HWP_COMPAT_CLASS: Readonly<Record<string, 'serif' | 'sans-serif' | 'monospace'>> = {",
        *[
            f"  '{f'HWP Compat {slug}'}': '{ {'serif': 'serif', 'sans': 'sans-serif', 'mono': 'monospace'}[info['class']] }',"
            for slug, info in sorted(built.items())
        ],
        "}",
    ]
    OUT_TS.write_text("\n".join(ts) + "\n", encoding="utf8")
    # advance tables for the converter's line fitting (hwpx-docx condense), as
    # runs of consecutive code points: [first, ...advances]
    metrics = {}
    for face, entry in data["faces"].items():
        table = entry["regular"]
        cps = sorted(int(cp, 16) for cp in table["advances"])
        runs, cur = [], None
        for cp in cps:
            adv = table["advances"][f"{cp:04X}"]
            if cur and cp == cur[0] + len(cur) - 1:
                cur.append(adv)
            else:
                cur = [cp, adv]
                runs.append(cur)
        metrics[face] = {"hangul": table.get("hangul", 1000), "runs": runs}
    OUT_METRICS.write_text(
        "// Generated by tools/hwp-compat/build-hwp-compat-fonts.py — do not edit.\n"
        "// Hancom advance widths (1/1000 em) from rhwp's tables (MIT), see tools/hwp-compat.\n"
        "/** face -> syllable advance and runs of [first code point, ...advances] */\n"
        "export const HWP_COMPAT_METRICS: Readonly<Record<string, { hangul: number; runs: readonly number[][] }>> =\n"
        f"  {json.dumps(metrics, ensure_ascii=False, separators=(',', ':'))}\n",
        encoding="utf8",
    )

    # the repository's formatter owns the TS/CSS layout; generated output must match it
    subprocess.run(
        ["npx", "prettier", "--write", str(OUT_TS), str(OUT_CSS), str(OUT_METRICS)],
        cwd=ROOT,
        check=True,
        capture_output=True,
    )
    total = sum((OUT_DIR / f"{s}.woff2").stat().st_size for s in built)
    print(f"{len(built)} faces for {len(names)} names, {total // 1024} KiB -> {OUT_DIR}")


if __name__ == "__main__":
    main()
