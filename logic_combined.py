"""
logic_combined.py
─────────────────
Thin wrapper that re-uses:
  • logic_equity.py  (must expose update_estimation(kotak, est))
  • logic_debt.py    (must expose process(est, kotak) -> output_file)

Exports:
  update_equity(kotak, est)  -> est   (in-place)
  update_debt(kotak, est)    -> file  (returns copy made by debt logic)
  update(kotak, est, section) -> file (section = 'equity'|'debt'|'both')
"""

from pathlib import Path
import importlib.util
import os

# ── Locate the two separate logic files ────────────────────────────────
EQ_PATH   = Path(__file__).with_name("logic_equity.py")
DEBT_PATH = Path(__file__).with_name("logic_debt.py")

if not EQ_PATH.exists():
    raise FileNotFoundError(f"Cannot find {EQ_PATH}")
if not DEBT_PATH.exists():
    raise FileNotFoundError(f"Cannot find {DEBT_PATH}")

_spec_eq = importlib.util.spec_from_file_location("logic_equity", EQ_PATH)
logic_equity = importlib.util.module_from_spec(_spec_eq)
_spec_eq.loader.exec_module(logic_equity)

_spec_debt = importlib.util.spec_from_file_location("logic_debt", DEBT_PATH)
logic_debt = importlib.util.module_from_spec(_spec_debt)
_spec_debt.loader.exec_module(logic_debt)

# ── Thin wrappers so callers get one consistent API ────────────────────
def update_equity(kotak_path: str | os.PathLike, est_path: str | os.PathLike) -> str:
    """Run the equity updater *in place*. Returns the (same) est_path."""
    logic_equity.update_estimation(kotak_path, est_path)
    return str(est_path)

def update_debt(kotak_path: str | os.PathLike, est_path: str | os.PathLike) -> str:
    """
    Run the debt updater. logic_debt.process() copies the file and returns
    the new filename. We forward that filename so callers know what to send.
    """
    return logic_debt.process(est_path, kotak_path)

# def update(kotak_path, est_path, section: str = "both") -> str:
#     """
#     section: 'equity' | 'debt' | 'both'
#     Returns the *final* path (might be a copy if debt update ran last).
#     """
#     section = section.lower()
#     if section not in {"equity", "debt", "both"}:
#         raise ValueError("section must be 'equity', 'debt', or 'both'")

#     base
#     out_path = est_path
#     if section in {"equity", "both"}:
#         out_path = update_equity(kotak_path, out_path)
#     if section in {"debt", "both"}:
#         out_path = update_debt(kotak_path, out_path)
#     return out_path
def update(kotak_path, est_path, section: str = "both") -> str:
    """
    section: 'equity' | 'debt' | 'both'
    Returns the *final* path (might be a copy if debt update ran last).
    """
    section = section.lower()
    if section not in {"equity", "debt", "both"}:
        raise ValueError("section must be 'equity', 'debt', or 'both'")

    out_path = est_path
    if section in {"equity", "both"}:
        out_path = update_equity(kotak_path, out_path)
    if section in {"debt", "both"}:
        out_path = update_debt(kotak_path, out_path)
    return out_path