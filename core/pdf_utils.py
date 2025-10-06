
# -*- coding: utf-8 -*-
from __future__ import annotations
import os, tempfile
from typing import List, Optional
from io import BytesIO

def try_docx_to_pdf(docx_bytes: bytes) -> Optional[bytes]:
    """Convierte DOCX a PDF usando docx2pdf si está disponible en el entorno (requiere MS Word en Windows/Mac)."""
    try:
        from docx2pdf import convert
    except Exception:
        return None
    with tempfile.TemporaryDirectory() as td:
        in_path = os.path.join(td, "in.docx")
        out_path = os.path.join(td, "out.pdf")
        with open(in_path, "wb") as f: f.write(docx_bytes)
        try:
            convert(in_path, out_path)
            with open(out_path, "rb") as f: return f.read()
        except Exception:
            return None

def merge_pdfs(pdf_bytes_list: List[bytes]) -> Optional[bytes]:
    try:
        from pypdf import PdfReader, PdfWriter
    except Exception:
        return None
    writer = PdfWriter()
    for b in pdf_bytes_list:
        try:
            reader = PdfReader(BytesIO(b))
            for page in reader.pages:
                writer.add_page(page)
        except Exception:
            continue
    out = BytesIO(); writer.write(out); return out.getvalue()

def add_text_watermark(pdf_bytes: bytes, text: str) -> Optional[bytes]:
    """Agrega marca de agua como texto simple (requiere reportlab y pypdf)."""
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        from pypdf import PdfReader, PdfWriter
    except Exception:
        return None
    from io import BytesIO
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont("Helvetica", 36)
    can.setFillGray(0.6, 0.4)
    can.saveState()
    can.translate(300, 400); can.rotate(45); can.drawCentredString(0, 0, text)
    can.restoreState(); can.save()
    packet.seek(0)
    watermark = PdfReader(packet)
    reader = PdfReader(BytesIO(pdf_bytes))
    writer = PdfWriter()
    for page in reader.pages:
        page.merge_page(watermark.pages[0])
        writer.add_page(page)
    out = BytesIO(); writer.write(out); return out.getvalue()

def sign_pdf_with_pfx(pdf_bytes: bytes, pfx_bytes: bytes, pfx_password: str) -> Optional[bytes]:
    """Firma digitalmente un PDF usando un PFX (requiere pyhanko)."""
    try:
        from pyhanko.sign import signers
        from pyhanko.sign.signers import pkcs12
    except Exception:
        return None
    try:
        signer = pkcs12.PKCS12Signer.load(pkcs12_data=pfx_bytes, passphrase=pfx_password.encode())
        signed_bytes = signers.sign_pdf(BytesIO(pdf_bytes), signer=signer)
        return signed_bytes
    except Exception:
        return None
