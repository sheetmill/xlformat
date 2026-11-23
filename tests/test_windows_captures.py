"""The Windows captures were made by the jobs committed beside them, from inputs written here.

Each ``tests/captures/windows/<group>.job.json`` is the manifest a job ran; every suite file of
the group names it and holds the job's result header (``job_header``). Writing the group's job
again from its ``job_id`` must give the same manifest, byte for byte, and inputs with the SHA-256
the result recorded.
"""

from __future__ import annotations

import gzip
import hashlib
import json
from pathlib import Path

import pytest

pytest.importorskip("openpyxl")

from capture.job import GROUPS, write
from capture.schema import read_rows

WINDOWS = Path(__file__).parent / "captures" / "windows"
JOBS = sorted(WINDOWS.glob("*.job.json")) if WINDOWS.exists() else []


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _header(group: str) -> dict[str, object]:
    """The job header of the group's first suite file, checked equal in every other one."""
    if group == "rows":  # the row-height layout file holds the rows group's header
        with gzip.open(WINDOWS / "row_height.layout.tsv.gz", "rt", encoding="utf-8") as f:
            line = next(x for x in f if x.startswith("# job_header: "))
        rows_header: dict[str, object] = json.loads(line.removeprefix("# job_header: "))
        return rows_header
    headers = []
    for name in GROUPS[group]:
        path = next(p for p in (WINDOWS / f"{name}.tsv", WINDOWS / f"{name}.tsv.gz") if p.exists())
        meta, _ = read_rows(path)
        assert meta["job"] == f"{group}.job.json"
        headers.append(meta["job_header"])
    assert len(set(headers)) == 1
    result: dict[str, object] = json.loads(headers[0])
    return result


@pytest.mark.parametrize("job", JOBS, ids=[p.name.removesuffix(".job.json") for p in JOBS])
def test_job_regenerates(job: Path, tmp_path: Path) -> None:
    group = job.name.removesuffix(".job.json")
    header = _header(group)
    committed = json.loads(job.read_text(encoding="utf-8"))
    assert header["job_id"] == committed["job_id"]
    assert header["manifest_sha256"] == _sha256(job)
    prefix = committed["purpose"].split("; ", 1)
    note = prefix[1] if len(prefix) == 2 else ""
    folder = write(group, str(committed["job_id"]), tmp_path, note)
    assert (folder / "job.json").read_bytes() == job.read_bytes()
    inputs = header["inputs"]
    assert isinstance(inputs, dict)
    assert sorted(inputs) == sorted(p.name for p in (folder / "files").iterdir())
    for name, entry in inputs.items():
        assert entry["sha256"] == _sha256(folder / "files" / name), name
