"""範例週報與 snapshot 的位置。"""
import json
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
SNAPSHOT_DIR = Path(__file__).parent / "fixtures" / "snapshots"


def sample_reports() -> list[Path]:
    return sorted(REPO_ROOT.glob("*週報.docx"))


def snapshot_path(docx_path: Path) -> Path:
    return SNAPSHOT_DIR / f"{docx_path.stem}.json"


def dump(snapshot: dict) -> str:
    return json.dumps(snapshot, ensure_ascii=False, indent=1)
