
# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import Dict
try:
    from docxcompose.composer import Composer
    from docx import Document
    HAS_DOCXCOMPOSE = True
except Exception:
    HAS_DOCXCOMPOSE = False
from io import BytesIO

def merge_documents_docx(named_docs: Dict[str, bytes]) -> bytes | None:
    if not HAS_DOCXCOMPOSE or not named_docs: return None
    names = sorted(named_docs.keys())
    master = Document(BytesIO(named_docs[names[0]]))
    composer = Composer(master)
    for name in names[1:]:
        composer.append(Document(BytesIO(named_docs[name])))
    out = BytesIO(); composer.save(out); return out.getvalue()
