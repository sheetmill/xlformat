"""The Windows captures were made by the jobs committed beside them, from inputs written here.

Each ``tests/captures/windows/<group>.job.json`` is the manifest a job ran; every suite file of
the group names it and holds the job's result header (``job_header``). Writing the group's job
again from its ``job_id`` must give the same manifest, byte for byte, and inputs with the SHA-256
the result recorded.

Writing a group's workbooks takes seconds, so the result (the manifest and each input's SHA-256)
is kept in pytest's cache under a key of everything that can change the bytes: the group, its
job_id and note, every file of ``capture/``, ``src/xlformat/`` and ``tests/captures/``, and the
versions of openpyxl, Python and zlib. A cold cache, as in CI, writes every group.
"""

from __future__ import annotations

import gzip
import hashlib
import json
import sys
import zlib
from functools import cache
from pathlib import Path

import pytest

pytest.importorskip("openpyxl")

import openpyxl
from capture.job import GROUPS, write
from capture.schema import read_header

from capture import faces_windows

ROOT = Path(__file__).parent.parent
WINDOWS = Path(__file__).parent / "captures" / "windows"
# the groups whose job writes one layout file instead of suite files
LAYOUTS = {"rows": "row_height.layout.tsv.gz", "builtin": "builtin_formats.layout.tsv"}
JOBS = sorted(WINDOWS.glob("*.job.json")) if WINDOWS.exists() else []


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _header(group: str) -> dict[str, object]:
    """The job header of the group's first suite file, checked equal in every other one."""
    if group.startswith("screen_"):  # the first of its layout files
        names = (f"{group}.layout.tsv", f"{group}_rules.layout.tsv", f"{group}_vertical.layout.tsv")
        group_layouts = [WINDOWS / name for name in names if (WINDOWS / name).exists()]
        assert group_layouts, group
        path = group_layouts[0]
        with path.open(encoding="utf-8") as f:
            line = next(x for x in f if x.startswith("# job_header: "))
        screen_header: dict[str, object] = json.loads(line.removeprefix("# job_header: "))
        for other in group_layouts[1:]:
            assert (
                f"# job_header: {json.dumps(screen_header, sort_keys=True, ensure_ascii=False)}\n"
                in (other.read_text(encoding="utf-8"))
            )
        return screen_header
    if group in faces_windows.JOB_GROUPS:  # each measure file's second line, the same in all
        faces_headers = set()
        for name in GROUPS[group]:
            with gzip.open(faces_windows.FACES_DIR / f"{name}.measure.tsv.gz", "rt") as f:
                f.readline()
                faces_headers.add(f.readline().removeprefix("# job_header: "))
        assert len(faces_headers) == 1
        faces_header: dict[str, object] = json.loads(faces_headers.pop())
        return faces_header
    if group in LAYOUTS:  # a layout file holds the header of its group
        path = WINDOWS / LAYOUTS[group]
        opener = gzip.open if path.suffix == ".gz" else open
        with opener(path, "rt", encoding="utf-8") as f:
            line = next(x for x in f if x.startswith("# job_header: "))
        layout_header: dict[str, object] = json.loads(line.removeprefix("# job_header: "))
        return layout_header
    headers = []
    for name in GROUPS[group]:
        path = next(p for p in (WINDOWS / f"{name}.tsv", WINDOWS / f"{name}.tsv.gz") if p.exists())
        meta = read_header(path)
        assert meta["job"] == f"{group}.job.json"
        headers.append(meta["job_header"])
    assert len(set(headers)) == 1
    result: dict[str, object] = json.loads(headers[0])
    return result


@cache
def _sources() -> str:
    """A digest of every file the written job can depend on, and of the tools' versions.

    The suites' cases are built by ``capture/`` (some from the captures' own rows), the workbooks
    with openpyxl; ``capture/build/`` holds local output only and is left out.
    """
    digest = hashlib.sha256()
    folders = (ROOT / "capture", ROOT / "src" / "xlformat", ROOT / "tests" / "captures")
    for folder in folders:
        for path in sorted(folder.rglob("*")):
            parts = path.relative_to(ROOT).parts
            if path.is_file() and parts[:2] != ("capture", "build") and "__pycache__" not in parts:
                digest.update("/".join(parts).encode() + b"\0" + path.read_bytes())
    for version in (openpyxl.__version__, sys.version, zlib.ZLIB_RUNTIME_VERSION):
        digest.update(version.encode() + b"\0")
    return digest.hexdigest()


def _written(
    group: str, job_id: str, note: str, tmp_path: Path, request: pytest.FixtureRequest
) -> tuple[str, dict[str, str]]:
    """The written job's manifest and its inputs' SHA-256, from pytest's cache when it has them."""
    store = getattr(request.config, "cache", None)  # absent under -p no:cacheprovider
    key = (
        "xlformat/jobs/"
        + hashlib.sha256(json.dumps([group, job_id, note, _sources()]).encode()).hexdigest()
    )
    if store is not None:
        kept = store.get(key, None)
        if kept is not None:
            return kept["manifest"], kept["inputs"]
    folder = write(group, job_id, tmp_path, note)
    manifest = (folder / "job.json").read_text(encoding="utf-8")
    inputs = {p.name: _sha256(p) for p in sorted((folder / "files").iterdir())}
    if store is not None:
        store.set(key, {"manifest": manifest, "inputs": inputs})
    return manifest, inputs


@pytest.mark.parametrize("job", JOBS, ids=[p.name.removesuffix(".job.json") for p in JOBS])
def test_job_regenerates(job: Path, tmp_path: Path, request: pytest.FixtureRequest) -> None:
    group = job.name.removesuffix(".job.json")
    header = _header(group)
    committed = json.loads(job.read_text(encoding="utf-8"))
    assert header["job_id"] == committed["job_id"]
    assert header["manifest_sha256"] == _sha256(job)
    prefix = committed["purpose"].split("; ", 1)
    note = prefix[1] if len(prefix) == 2 else ""
    manifest, written = _written(group, str(committed["job_id"]), note, tmp_path, request)
    assert manifest.encode("utf-8") == job.read_bytes()
    inputs = header["inputs"]
    assert isinstance(inputs, dict)
    assert sorted(inputs) == sorted(written)
    for name, entry in inputs.items():
        assert entry["sha256"] == written[name], name
