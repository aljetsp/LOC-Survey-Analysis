"""Interactive command line: raw export in, analysis workbook out.

    python -m survey_tool.cli build raw.xlsx --reference city_reference.csv

The prompts exist because three inputs change every year and cannot be inferred
safely: current populations, the quintile breakpoints, and the region for any
city not already on file.
"""
from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

from .pipeline import build as run_build
from .pipeline import make_codebook
from .profile import profile_workbook, read_codebook
from .reference import (DEFAULT_BREAKPOINTS, REGION_NAMES, CityReference,
                        normalize_city, quintile_for)


def _ask(prompt, default=None):
    suffix = f" [{default}]" if default is not None else ""
    try:
        val = input(f"{prompt}{suffix}: ").strip()
    except EOFError:
        return default
    return val or default


def _yes(prompt, default=True):
    d = "Y/n" if default else "y/N"
    v = _ask(f"{prompt} ({d})", "")
    if not v:
        return default
    return v.lower().startswith("y")


# --------------------------------------------------------------------------
def update_populations(ref: CityReference):
    """Load a fresh population file, or edit individual cities."""
    print("\n-- Populations --")
    print(f"    {len(ref.cities)} cities on file.")
    path = _ask("    Path to updated population CSV (columns: city,population), "
                "blank to skip", "")
    if path:
        p = Path(path).expanduser()
        if not p.exists():
            print(f"    ! {p} not found -- skipping.")
        else:
            changed = added = 0
            with open(p, newline="", encoding="utf-8-sig") as fh:
                for row in csv.DictReader(fh):
                    name = normalize_city(row.get("city", ""))
                    raw = (row.get("population") or "").replace(",", "").strip()
                    if not name or not raw:
                        continue
                    pop = float(raw)
                    existing = ref.cities.get(name)
                    if existing is None:
                        added += 1
                    elif existing.population != pop:
                        changed += 1
                    ref.upsert(name, population=pop)
            print(f"    {changed} population(s) updated, {added} city(ies) added.")

    while _yes("    Edit an individual city's population?", False):
        name = _ask("      City")
        if not name:
            break
        city, conf = ref.match(name)
        if city is None:
            print("      Not on file -- it will be added.")
            city = ref.upsert(name)
        pop = _ask(f"      Population for {city.name}",
                   "" if city.population is None else int(city.population))
        if pop not in (None, ""):
            ref.upsert(city.name, population=float(str(pop).replace(",", "")))


def confirm_breakpoints(ref: CityReference):
    """Quintiles are derived from population bands, and the bands move."""
    print("\n-- Quintile breakpoints --")
    print("    A city falls in quintile 1 below the first value, 2 below the")
    print("    second, and so on; quintile 5 is everything above the last.")
    print(f"    Current: {', '.join(f'{b:,}' for b in ref.breakpoints)}")
    print(f"    (Derived from the 2025/2026 workbooks; the Q4/Q5 line moved "
          f"between those two years, so confirm it against this year's list.)")
    if _yes("    Change them?", False):
        vals = _ask("      Four ascending numbers, comma separated",
                    ",".join(str(b) for b in ref.breakpoints))
        try:
            bps = [float(v.strip().replace(",", "")) for v in str(vals).split(",")]
            if len(bps) != 4 or bps != sorted(bps):
                raise ValueError
            ref.breakpoints = bps
        except ValueError:
            print("      ! Need exactly four ascending numbers -- keeping current values.")

    counts = {q: 0 for q in range(1, 6)}
    edge = []
    for c in ref.cities.values():
        if c.population is None:
            continue
        q = quintile_for(c.population, ref.breakpoints)
        counts[q] += 1
        for b in ref.breakpoints:
            if 0 < abs(c.population - b) / b <= 0.03:
                edge.append((c.name, int(c.population), q))
    print("    Resulting distribution: "
          + ", ".join(f"Q{q}={n}" for q, n in counts.items()))
    if edge:
        print(f"    {len(edge)} city(ies) sit within 3% of a boundary:")
        for name, pop, q in sorted(edge, key=lambda x: x[1])[:10]:
            print(f"      {name:22s} {pop:>8,}  -> Q{q}")


def resolve_regions(ref: CityReference, needed: set[str]):
    """Every city in the response set needs a region before analysis can run."""
    missing = [n for n in sorted(needed)
               if n in ref.cities and ref.cities[n].region_code is None]
    print("\n-- Regions --")
    if not missing:
        print("    All responding cities have a region on file.")
        return
    print(f"    {len(missing)} responding city(ies) have no region:")
    for code, name in sorted(REGION_NAMES.items()):
        print(f"      {code:2d} = {name}")
    for name in missing:
        while True:
            v = _ask(f"      Region code for {name} (1-12, blank to drop the city)", "")
            if v in (None, ""):
                print(f"        {name} will be excluded from the analysis.")
                break
            try:
                code = int(v)
                if 1 <= code <= 12:
                    ref.upsert(name, region_code=code)
                    break
            except ValueError:
                pass
            print("        Enter a number from 1 to 12.")


def resolve_cities(raw_path, codebook_path, ref: CityReference):
    """Confirm approximate matches and place unknown city names."""
    specs, meta = read_codebook(codebook_path)
    import openpyxl
    ws = openpyxl.load_workbook(raw_path, data_only=True)[meta["sheet"]]
    city_spec = next(s for s in specs if s.role == "city")

    raw_names, overrides = [], {}
    for r in range(meta["first_data_row"], meta["last_data_row"] + 1):
        v = ws.cell(r, city_spec.index).value
        if v not in (None, ""):
            raw_names.append(str(v))

    resolved, fuzzy, unknown = set(), [], []
    for rn in raw_names:
        city, conf = ref.match(rn)
        if conf == "exact":
            resolved.add(city.name)
        elif conf == "fuzzy":
            fuzzy.append((rn, city.name))
        else:
            unknown.append(rn)

    print("\n-- City names --")
    print(f"    {len(raw_names)} responses; {len(resolved)} matched exactly.")

    for rn, guess in fuzzy:
        if _yes(f"    '{rn}' -> {guess}?  Accept?", True):
            overrides[rn] = guess
            resolved.add(guess)
        else:
            unknown.append(rn)

    for rn in unknown:
        print(f"    '{rn}' is not on file.")
        v = _ask("      Correct city name (blank to drop this response)", "")
        if not v:
            continue
        city, conf = ref.match(v)
        if city is None:
            city = ref.upsert(v)
            print(f"      Added {city.name} to the reference.")
            pop = _ask(f"      Population for {city.name}", "")
            if pop:
                ref.upsert(city.name, population=float(str(pop).replace(",", "")))
        overrides[rn] = city.name
        resolved.add(city.name)

    return overrides, resolved


# --------------------------------------------------------------------------
def cmd_build(args):
    raw = Path(args.raw)
    out_dir = Path(args.outdir)
    out_dir.mkdir(parents=True, exist_ok=True)
    codebook = out_dir / f"{raw.stem}.codebook.json"

    print(f"Reading {raw.name} ...")
    specs, sheet, first, last = make_codebook(raw, codebook, sheet=args.sheet)
    n_q = sum(1 for s in specs if s.role == "question" and s.kind != "skip")
    n_rev = sum(1 for s in specs if s.note.startswith("REVIEW"))
    print(f"    sheet '{sheet}', rows {first}-{last}, {n_q} usable question columns.")
    print(f"    codebook written to {codebook}")
    if n_rev:
        print(f"    {n_rev} column(s) marked REVIEW -- open the codebook and check "
              f"the response codes before publishing.")

    ref = CityReference.load(args.reference)
    if not args.no_prompt:
        overrides, responding = resolve_cities(raw, codebook, ref)
        update_populations(ref)
        confirm_breakpoints(ref)
        resolve_regions(ref, responding)
        ref.save(args.reference)
        print(f"\n    Reference saved to {args.reference}")
    else:
        overrides = {}

    out = out_dir / f"{raw.stem}.analysis.xlsx"
    awb, res = run_build(raw, codebook, args.reference, out,
                         title=args.title or raw.stem, city_overrides=overrides)
    print(f"\nWrote {out}")
    print(f"    {res.n} respondents, {len(res.columns)} analysis columns")
    print(f"    {len(res.issues)} entry(ies) in the Cleaning Log")
    print("\nNext: recalculate the workbook so cached values are populated:")
    print(f"    python -m survey_tool.cli recalc {out}")
    return 0


def cmd_profile(args):
    specs, sheet, (first, last) = profile_workbook(args.raw, sheet=args.sheet)
    print(f"sheet '{sheet}', data rows {first}-{last}")
    for s in specs:
        if s.role == "metadata":
            continue
        flag = "  <-- REVIEW" if s.note.startswith("REVIEW") else ""
        print(f"  col {s.index:4d}  {s.role:9s} {s.kind:14s} {s.text[:56]}{flag}")
    return 0


def cmd_recalc(args):
    """Populate cached values via LibreOffice and report any formula errors."""
    import json
    import subprocess
    script = Path(args.script).expanduser()
    if not script.exists():
        print(f"! recalc script not found at {script}")
        print("  Point --script at the xlsx skill's scripts/recalc.py, or open and")
        print("  re-save the workbook in Excel, which recalculates on open.")
        return 1
    proc = subprocess.run([sys.executable, str(script), args.workbook, str(args.timeout)],
                          capture_output=True, text=True)
    try:
        report = json.loads(proc.stdout)
    except json.JSONDecodeError:
        print(proc.stdout or proc.stderr)
        return 1
    if "error" in report:
        print(f"! nothing was recalculated: {report['error']}")
        return 1
    print(f"{report['total_formulas']:,} formulas, {report['total_errors']} error(s)")
    for kind, detail in report.get("error_summary", {}).items():
        print(f"  {kind}: {detail['count']} -- e.g. {', '.join(detail['locations'][:5])}")
    return 0 if report["total_errors"] == 0 else 1


def main(argv=None):
    p = argparse.ArgumentParser(prog="survey_tool", description=__doc__)
    sub = p.add_subparsers(dest="cmd", required=True)

    b = sub.add_parser("build", help="clean a raw export and write the analysis workbook")
    b.add_argument("raw")
    b.add_argument("--reference", default="city_reference.csv")
    b.add_argument("--outdir", default="output")
    b.add_argument("--sheet", default=None)
    b.add_argument("--title", default=None)
    b.add_argument("--no-prompt", action="store_true",
                   help="skip the interactive prompts and use the reference as-is")
    b.set_defaults(func=cmd_build)

    q = sub.add_parser("profile", help="show how each column would be treated")
    q.add_argument("raw")
    q.add_argument("--sheet", default=None)
    q.set_defaults(func=cmd_profile)

    rc = sub.add_parser("recalc", help="recalculate a built workbook and check for errors")
    rc.add_argument("workbook")
    rc.add_argument("--script", default="/mnt/skills/public/xlsx/scripts/recalc.py")
    rc.add_argument("--timeout", type=int, default=180)
    rc.set_defaults(func=cmd_recalc)

    args = p.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
