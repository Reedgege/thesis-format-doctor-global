"""IEEE conference templates are saved as Strict OOXML; python-docx cannot open
them until namespaces are rewritten to Transitional.

This test takes a real python-docx document, converts it to Strict OOXML via the
reverse of the production mapping, then verifies the engine's normaliser makes it
readable again (and that a direct open of the strict copy fails, proving the fix
is necessary).
"""

from __future__ import annotations

import io
import os
import sys
import zipfile

from docx import Document

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.docx_reader import (
    _normalize_strict_ooxml,
    _STRICT_TO_TRANSITIONAL,
    read_docx,
)


def _make_real_docx(path):
    d = Document()
    d.add_paragraph("Plain body paragraph for IEEE two-column test.")
    d.add_paragraph("References")
    d.add_paragraph('J. A. Smith, "Title," J. Name, vol. 1, no. 2, pp. 3-4, 2020.')
    d.save(str(path))


def _to_strict_docx(src_path, dst_path):
    """Reverse the production mapping (per XML part) to make a Strict-OOXML copy."""
    reverse = {trans: strict for strict, trans in _STRICT_TO_TRANSITIONAL.items()}
    rev_sorted = sorted(reverse.items(), key=lambda kv: len(kv[0]), reverse=True)
    with open(src_path, "rb") as fh:
        raw = fh.read()
    src = io.BytesIO(raw)
    out = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            lower = item.filename.lower()
            if lower.endswith(".xml") or lower.endswith(".rels"):
                for trans, strict in rev_sorted:
                    if trans in data:
                        data = data.replace(trans, strict)
            zout.writestr(item, data)
    out.seek(0)
    with open(dst_path, "wb") as fh:
        fh.write(out.read())


def test_transitional_docx_is_noop(tmp_path):
    p = tmp_path / "normal.docx"
    _make_real_docx(p)
    assert _normalize_strict_ooxml(str(p)) is None


def test_read_docx_reads_transitional(tmp_path):
    p = tmp_path / "normal.docx"
    _make_real_docx(p)
    prof = read_docx(str(p))
    assert prof is not None


def test_strict_docx_direct_open_fails(tmp_path):
    real = tmp_path / "real.docx"
    _make_real_docx(real)
    strict = tmp_path / "strict.docx"
    _to_strict_docx(real, strict)
    # Sanity: the strict copy cannot be opened directly by python-docx.
    try:
        Document(str(strict))
        opened = True
    except Exception:
        opened = False
    assert opened is False


def test_strict_docx_becomes_readable_via_engine(tmp_path):
    real = tmp_path / "real.docx"
    _make_real_docx(real)
    strict = tmp_path / "strict.docx"
    _to_strict_docx(real, strict)

    # The engine must open it via normalisation.
    prof = read_docx(str(strict))
    assert prof is not None

    # Normalisation returns a stream python-docx can parse, with content intact.
    norm = _normalize_strict_ooxml(str(strict))
    assert norm is not None
    doc = Document(norm)
    text = "\n".join(par.text for par in doc.paragraphs)
    assert "Plain body paragraph for IEEE two-column test." in text
