"""
Resource Path Resolver
Resolves paths for internal assets, templates, and samples in both dev mode and PyInstaller bundles.
"""

import sys
import os


def get_base_dir() -> str:
    """Returns the base project directory."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    # Root of x2t_desktop
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_resource_path(relative_path: str) -> str:
    """
    Returns absolute path to a resource file.
    Works whether running as Python script or packaged with PyInstaller.
    """
    base = get_base_dir()
    target = os.path.join(base, relative_path)
    if os.path.exists(target):
        return target
    # Fallback check one level up (workspace root)
    workspace_root = os.path.dirname(base)
    fallback = os.path.join(workspace_root, relative_path)
    if os.path.exists(fallback):
        return fallback
    return target


def ensure_sample_file(sample_key: str) -> str:
    """
    Ensures sample Excel files exist and returns the path.
    """
    file_map = {
        "sales": ("sample_data", "sample_sales.xlsx"),
        "purchase": ("sample_data", "sample_purchase.xlsx"),
        "bank": ("sample_data", "sample_bank_statement.xlsx"),
        "vertical": ("sample_data", "sample_vertical_journal.xlsx")
    }
    if sample_key not in file_map:
        return ""

    folder, filename = file_map[sample_key]
    path = get_resource_path(os.path.join(folder, filename))
    if os.path.exists(path):
        return path

    # If not found, generate dynamically
    from templates.sample_generator import generate_all_samples
    target_dir = os.path.dirname(path)
    os.makedirs(target_dir, exist_ok=True)
    created = generate_all_samples(target_dir)
    return created.get(sample_key, path)
