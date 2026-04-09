from __future__ import annotations

import shutil
import sys
from pathlib import Path


SOURCE_ROOT = Path(__file__).resolve().parents[2]


def resource_base_dir() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS"))
    return SOURCE_ROOT


def runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return SOURCE_ROOT


def config_path() -> Path:
    return runtime_base_dir() / "config.json"


def data_dir() -> Path:
    return runtime_base_dir() / "data"


def exports_dir() -> Path:
    return runtime_base_dir() / "exports"


def sample_project_path() -> Path:
    return data_dir() / "sample_project.json"


def ensure_runtime_structure() -> None:
    data_dir().mkdir(parents=True, exist_ok=True)
    exports_dir().mkdir(parents=True, exist_ok=True)

    source_config = resource_base_dir() / "config.json"
    target_config = config_path()
    if source_config.exists() and not target_config.exists():
        shutil.copy2(source_config, target_config)

    source_sample = resource_base_dir() / "data" / "sample_project.json"
    target_sample = sample_project_path()
    if source_sample.exists() and not target_sample.exists():
        shutil.copy2(source_sample, target_sample)
