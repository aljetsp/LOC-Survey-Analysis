"""One-call pipeline: raw export in, analysis workbook out."""
from __future__ import annotations

from pathlib import Path

from .build import AnalysisWorkbook
from .clean import clean
from .profile import profile_workbook, read_codebook, write_codebook
from .reference import CityReference
from .validate import write_issue_log, write_validation


def make_codebook(raw_path, codebook_path, sheet=None):
    """Step 1: scan the raw export and write a reviewable codebook."""
    specs, sheet, (first, last) = profile_workbook(raw_path, sheet=sheet)
    write_codebook(specs, codebook_path,
                   meta={"source": str(raw_path), "sheet": sheet,
                         "first_data_row": first, "last_data_row": last})
    return specs, sheet, first, last


def build(raw_path, codebook_path, reference_path, out_path,
          title=None, city_overrides=None):
    """Step 2: clean against the reviewed codebook and write the workbook."""
    specs, meta = read_codebook(codebook_path)
    ref = CityReference.load(reference_path)
    result = clean(raw_path, specs, ref, meta["sheet"],
                   meta["first_data_row"], meta["last_data_row"],
                   city_overrides=city_overrides)

    awb = AnalysisWorkbook(result, ref, title=title or Path(raw_path).stem)
    awb.write_data()
    analysis = awb.write_analysis()
    write_validation(awb, analysis)
    write_issue_log(awb)
    awb.wb.save(out_path)
    return awb, result
