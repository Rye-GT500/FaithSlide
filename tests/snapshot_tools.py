"""範例週報與 snapshot 的位置、snapshot 產生方式。"""
import json
from pathlib import Path

from tests import legacy_runner
from tests.slide_extract import extract_slides

REPO_ROOT = Path(__file__).parent.parent
SNAPSHOT_DIR = Path(__file__).parent / "fixtures" / "snapshots"


def sample_reports() -> list[Path]:
    return sorted(REPO_ROOT.glob("*週報.docx"))


def snapshot_path(docx_path: Path) -> Path:
    return SNAPSHOT_DIR / f"{docx_path.stem}.json"


def build_snapshot(docx_path: Path, out_pptx: Path, cache: legacy_runner.VerseCache) -> dict:
    messages = legacy_runner.run(docx_path, out_pptx, cache)
    slides = extract_slides(out_pptx) if out_pptx.exists() else None
    return {"source": docx_path.name, "messages": messages, "slides": slides}


def dump(snapshot: dict) -> str:
    return json.dumps(snapshot, ensure_ascii=False, indent=1)
