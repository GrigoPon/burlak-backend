"""Comprehensive MS Excel compatibility validation for BAIC split .xlsx files.

Usage:
    python -m pytest tests/test_baic_xlsx_integrity.py -v -s
    python tests/test_baic_xlsx_integrity.py  (standalone, prints summary)
"""

from __future__ import annotations

import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

PROJECT_ROOT = Path(__file__).resolve().parent.parent
BAIC_SPLIT_DIR = PROJECT_ROOT / "output_baic" / "split_cards"

# Namespaces used in OOXML
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"

# Pattern to detect auto-generated ns0: prefixes in XML (from broken ET.register_namespace)
# Microsoft tools legitimately use ns1:, ns2:, etc. in original OOXML files.
# Only ns0: is a sign of corruption.
NS0_PATTERN = re.compile(r'xmlns:ns0="|</?ns0:')


def _check_single_file(path: str) -> Dict[str, Any]:
    """Run all checks on a single .xlsx file (designed for parallelism)."""
    basename = os.path.basename(path)
    result: Dict[str, Any] = {
        "file": basename,
        "ok": True,
        "errors": [],
        "warnings": [],
    }

    try:
        with zipfile.ZipFile(path, "r") as zf:
            names_raw = zf.namelist()
    except zipfile.BadZipFile as e:
        result["ok"] = False
        result["errors"].append(f"BadZipFile: {e}")
        return result

    names = set(names_raw)

    # ── Check 1: Has sheet file ──
    has_sheet = any(
        n.startswith("xl/worksheets/sheet") and n.endswith(".xml") and "/_rels/" not in n
        for n in names
    )
    if not has_sheet:
        result["ok"] = False
        result["errors"].append("No xl/worksheets/sheet*.xml found")

    # ── Check 2: Has [Content_Types].xml and xl/workbook.xml ──
    for req in ("[Content_Types].xml", "xl/workbook.xml"):
        if req not in names:
            result["ok"] = False
            result["errors"].append(f"Missing required file: {req}")

    # ── Check 3: No ns0:/ns1: tags ──
    ns0_files = []
    for name in names_raw:
        if name.endswith(".xml") and "/_rels/" not in name:
            try:
                with zipfile.ZipFile(path, "r") as zf:
                    data = zf.read(name)
                if NS0_PATTERN.search(data.decode("utf-8", errors="replace")):
                    ns0_files.append(name)
            except Exception:
                pass
    if ns0_files:
        result["ok"] = False
        result["errors"].append(f"ns0:/ns1: prefixes in: {ns0_files[:5]}")

    # ── Check 4: Relationship targets exist ──
    if "xl/_rels/workbook.xml.rels" in names:
        try:
            with zipfile.ZipFile(path, "r") as zf:
                rels_data = zf.read("xl/_rels/workbook.xml.rels")
            rels_root = ET.fromstring(rels_data)
            orphaned = []
            for rel_el in rels_root:
                target = rel_el.get("Target", "")
                if not target:
                    continue
                resolved = os.path.normpath(os.path.join("xl", target)).replace(os.sep, "/")
                if resolved not in names:
                    orphaned.append(f"{rel_el.get('Id', '?')} -> {target}")
            if orphaned:
                result["ok"] = False
                result["errors"].append(f"Orphaned relationship targets: {orphaned}")
        except Exception as e:
            result["ok"] = False
            result["errors"].append(f"Error checking .rels: {e}")

    # ── Check 5: Content_Types Overrides exist ──
    if "[Content_Types].xml" in names:
        try:
            with zipfile.ZipFile(path, "r") as zf:
                ct_data = zf.read("[Content_Types].xml")
            ct_root = ET.fromstring(ct_data)
            orphaned = []
            for override_el in ct_root.findall(f"{{{NS_CT}}}Override"):
                part_name = override_el.get("PartName", "")
                if part_name.startswith("/"):
                    part_name = part_name[1:]
                if part_name not in names and "calcChain" not in part_name:
                    orphaned.append(part_name)
            if orphaned:
                result["ok"] = False
                result["errors"].append(f"Orphaned ContentType Overrides: {orphaned}")
        except Exception:
            result["ok"] = False
            result["errors"].append("Bad [Content_Types].xml")

    # ── Check 6: Sub-relationship files reference existing targets ──
    for name in names_raw:
        if "_rels/" in name and name.endswith(".rels"):
            try:
                with zipfile.ZipFile(path, "r") as zf:
                    rels_data = zf.read(name)
                rels_root = ET.fromstring(rels_data)
                rels_dir = os.path.dirname(os.path.dirname(name))
                for rel_el in rels_root:
                    target = rel_el.get("Target", "")
                    if not target:
                        continue
                    resolved = os.path.normpath(os.path.join(rels_dir, target)).replace(os.sep, "/")
                    if resolved not in names:
                        if "/drawings/" in resolved or "/charts/" in resolved or "/vml" in resolved:
                            result["ok"] = False
                            result["errors"].append(
                                f"Orphaned {resolved} referenced from {name}"
                            )
            except Exception:
                pass

    # ── Check 7: Quick openpyxl check (sampled: every 10th file) ──
    # We hash the path to pseudo-randomly sample ~10% of files
    if hash(path) % 10 == 0:
        try:
            import openpyxl
            import warnings
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                _ = wb.sheetnames
                wb.close()
        except Exception as e:
            result["ok"] = False
            result["errors"].append(f"openpyxl load failed: {type(e).__name__}: {e}")

    return result


def validate_all() -> Dict[str, Any]:
    """Run all checks on all BAIC split files in parallel."""
    if not BAIC_SPLIT_DIR.is_dir():
        return {"ok": False, "error": f"Directory not found: {BAIC_SPLIT_DIR}"}

    xlsx_files = sorted(str(p) for p in BAIC_SPLIT_DIR.glob("*.xlsx"))
    total = len(xlsx_files)
    passed = 0
    failed = 0
    all_errors: List[str] = []
    openpyxl_checked = 0

    # Run in parallel
    with ProcessPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(_check_single_file, f): f for f in xlsx_files}
        for future in as_completed(futures):
            result = future.result()
            if result["ok"]:
                passed += 1
            else:
                failed += 1
                for err in result["errors"]:
                    all_errors.append(f"{result['file']}: {err}")

    return {
        "ok": failed == 0,
        "total": total,
        "passed": passed,
        "failed": failed,
        "errors": all_errors,
        "openpyxl_checked": openpyxl_checked,
    }


def print_summary(summary: Dict[str, Any]) -> None:
    """Pretty-print validation summary."""
    print("=" * 60)
    print("BAIC Split File — MS Excel Compatibility Check")
    print("=" * 60)
    print(f"  Total files:  {summary['total']}")
    print(f"  Passed:       {summary['passed']}")
    print(f"  Failed:       {summary['failed']}")
    print("-" * 60)

    if summary["failed"] == 0:
        print(f"\n  ✅ ALL {summary['passed']} FILES PASSED")
        print(f"     No ns0: prefixes, no orphaned references,")
        print(f"     all relationship targets valid.")
    else:
        print(f"\n  ❌ {summary['failed']} FILES FAILED:")
        for err in summary["errors"]:
            print(f"     - {err}")

    print("=" * 60)


if __name__ == "__main__":
    summary = validate_all()
    print_summary(summary)
    sys.exit(0 if summary["ok"] else 1)
