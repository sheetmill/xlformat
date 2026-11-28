"""Every Windows capture carrying a job's ``manifest_sha256`` came from a committed manifest.

Each file under ``tests/captures/windows/`` holds its job's result header, and with it the
SHA-256 of the manifest the job ran (``manifest_sha256``). That hash must be the SHA-256 of one of
the committed ``<group>.job.json`` files, and of the very one a file names on its ``job:`` line.
Needs no development tools, so the sdist's run checks it too.
"""

from __future__ import annotations

import gzip
import hashlib
import re
from pathlib import Path

import pytest

WINDOWS = Path(__file__).parent / "captures" / "windows"
MANIFESTS = {
    hashlib.sha256(p.read_bytes()).hexdigest(): p.name for p in sorted(WINDOWS.glob("*.job.json"))
}
_HASH = re.compile(r'"manifest_sha256": "([0-9a-f]{64})"')
_JOB = re.compile(r"\bjob: ([\w.]+\.job\.json)")


def _text(path: Path) -> str:
    raw = path.read_bytes()
    return (gzip.decompress(raw) if path.suffix == ".gz" else raw).decode("utf-8", "replace")


CARRIERS = [
    p
    for p in sorted(WINDOWS.rglob("*"))
    if p.is_file() and not p.name.endswith(".job.json") and _HASH.search(_text(p))
]


def test_the_captures_carry_manifest_hashes() -> None:
    assert len(CARRIERS) >= 756


@pytest.mark.parametrize("path", CARRIERS, ids=lambda p: p.relative_to(WINDOWS).as_posix())
def test_a_capture_names_a_committed_manifest(path: Path) -> None:
    text = _text(path)
    hashes = set(_HASH.findall(text))
    assert len(hashes) == 1, hashes
    (found,) = hashes
    assert found in MANIFESTS, f"{found[:12]} is no committed manifest's SHA-256"
    named = _JOB.search(text)
    if named:
        assert MANIFESTS[found] == named[1], (MANIFESTS[found], named[1])
