"""City reference data: population, quintile band, and LOC region."""
from __future__ import annotations

import csv
import difflib
import re
from dataclasses import dataclass, field
from pathlib import Path

REGION_NAMES = {
    1: "N. Coast", 2: "Metro", 3: "N. Willamette", 4: "S. Willamette",
    5: "C. Coast", 6: "S. Coast", 7: "S. Oregon", 8: "Gorge",
    9: "C. Oregon", 10: "SC Oregon", 11: "NE Oregon", 12: "E. Oregon",
}

# Population bands inferred from the 2025/2026 workbooks. Each file is internally
# consistent, but the Q4/Q5 line moved between the 2025 and 2026 surveys
# (Cottage Grove: 10,879 -> Q5 in 2025, 10,909 -> Q4 in 2026), so no single set of
# breakpoints reproduces every file. These reproduce 426 of 427 assignments and are
# a starting point only -- confirm_breakpoints() asks the user to confirm each year.
DEFAULT_BREAKPOINTS = [500, 1_400, 3_300, 10_700]

QUINTILE_LABELS = {
    1: "1st Quintile", 2: "2nd Quintile", 3: "3rd Quintile",
    4: "4th Quintile", 5: "5th Quintile",
}


def quintile_for(population: float, breakpoints=None) -> int:
    """Return 1-5 for a population using ascending band breakpoints."""
    bps = list(breakpoints or DEFAULT_BREAKPOINTS)
    if len(bps) != 4:
        raise ValueError("expected exactly 4 breakpoints")
    if population is None or population == "":
        raise ValueError("population required to assign a quintile")
    pop = float(population)
    for i, bp in enumerate(bps):
        if pop < bp:
            return i + 1
    return 5


def normalize_city(raw: str) -> str:
    """Fold the spelling variants seen in Qualtrics free-text city fields."""
    if raw is None:
        return ""
    s = str(raw).strip()
    s = re.sub(r"[\u2018\u2019\u201c\u201d]", "'", s)
    # Respondents type this prefix by hand, so it arrives mangled in both tokens:
    # "Ctiy of", "City or", "Cit fo". Match each loosely rather than exactly, but
    # only when something remains afterwards -- "City o" alone is not a city.
    parts = s.split()
    if len(parts) >= 3:
        near_city = difflib.SequenceMatcher(None, parts[0].lower(), "city").ratio() >= 0.7
        near_town = difflib.SequenceMatcher(None, parts[0].lower(), "town").ratio() >= 0.7
        near_of = difflib.SequenceMatcher(None, parts[1].lower(), "of").ratio() >= 0.5
        if (near_city or near_town) and near_of:
            s = " ".join(parts[2:])
    s = re.sub(r"(?i)^(the\s+)?city\s+of\s+", "", s)
    s = re.sub(r"(?i)^town\s+of\s+", "", s)

    # Respondents often append their department: "The Dalles Public Works".
    s = re.sub(r"(?i)\s+(public\s+works|city\s+hall|water\s+dept\.?|"
               r"water\s+department|finance\s+dept\.?|administration)\s*$", "", s)
    s = re.sub(r"(?i),?\s*(oregon|ore\.?|or)$", "", s)
    s = re.sub(r"\s+", " ", s).strip(" .,")
    if not s:
        return ""
    # Title-case only if the input carries no deliberate mixed case (e.g. "HEPPNER")
    if s.isupper() or s.islower():
        s = " ".join(w.capitalize() for w in s.split())
    s = re.sub(r"(?i)\bMc([a-z])", lambda m: "Mc" + m.group(1).upper(), s)
    s = re.sub(r"(?i)^st\.?\s+", "St. ", s)
    s = re.sub(r"(?i)^mt\.?\s+", "Mt. ", s)
    return s


@dataclass
class City:
    name: str
    population: float | None = None
    region_code: int | None = None

    @property
    def region_name(self) -> str:
        return REGION_NAMES.get(self.region_code, "")


@dataclass
class CityReference:
    """The master lookup. Populations and regions are supplied by the user;
    quintiles are derived from populations so the two can never disagree."""

    cities: dict[str, City] = field(default_factory=dict)
    breakpoints: list[int] = field(default_factory=lambda: list(DEFAULT_BREAKPOINTS))

    # ---------- io ----------
    @classmethod
    def load(cls, path: str | Path) -> "CityReference":
        ref = cls()
        with open(path, newline="", encoding="utf-8-sig") as fh:
            for row in csv.DictReader(fh):
                name = normalize_city(row.get("city", ""))
                if not name:
                    continue
                pop = row.get("population") or ""
                reg = row.get("region_code") or ""
                ref.cities[name] = City(
                    name=name,
                    population=float(pop) if str(pop).strip() != "" else None,
                    region_code=int(float(reg)) if str(reg).strip() != "" else None,
                )
        return ref

    def save(self, path: str | Path) -> None:
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["city", "population", "region_code", "region_name", "quintile"])
            for name in sorted(self.cities):
                c = self.cities[name]
                q = ""
                if c.population is not None:
                    q = quintile_for(c.population, self.breakpoints)
                w.writerow([c.name,
                            "" if c.population is None else int(c.population),
                            "" if c.region_code is None else c.region_code,
                            c.region_name, q])

    # ---------- lookup ----------
    def match(self, raw_name: str, cutoff: float = 0.84):
        """Resolve a survey-entered city name.

        Returns (City, confidence) where confidence is 'exact', 'fuzzy', or None.
        Fuzzy matches are surfaced to the user for confirmation rather than
        applied silently -- a wrong match silently reassigns a city's quintile.
        """
        name = normalize_city(raw_name)
        if not name:
            return None, None
        if name in self.cities:
            return self.cities[name], "exact"
        close = difflib.get_close_matches(name, list(self.cities), n=1, cutoff=cutoff)
        if close:
            return self.cities[close[0]], "fuzzy"
        return None, None

    def upsert(self, name: str, population=None, region_code=None) -> City:
        key = normalize_city(name)
        c = self.cities.get(key) or City(name=key)
        if population is not None:
            c.population = float(population)
        if region_code is not None:
            c.region_code = int(region_code)
        self.cities[key] = c
        return c

    def quintile(self, city: City) -> int | None:
        if city.population is None:
            return None
        return quintile_for(city.population, self.breakpoints)

    def missing_fields(self) -> list[str]:
        out = []
        for name in sorted(self.cities):
            c = self.cities[name]
            gaps = []
            if c.population is None:
                gaps.append("population")
            if c.region_code is None:
                gaps.append("region")
            if gaps:
                out.append(f"{name}: missing {', '.join(gaps)}")
        return out
