
# -*- coding: utf-8 -*-
from __future__ import annotations
import io
import pandas as pd
import streamlit as st
from docx import Document
from datetime import date

from core.backend import guess_mapping, prepare_dataframe
from core.routing import load_routing_yaml
from core.funcionalidades import generate_letters_per_group, build_index_sheet, make_zip
from core.merge import merge_documents_docx
from core.pdf_utils import try_docx_to_pdf, merge_pdfs, add_text_watermark, sign_pdf_with_pfx
from core.quality import compute_missing_summary, compute_duplicates_by_actor, compute_date_ranges_by_actor

def run_app() -> None:
    st.set_page_config(page_title="Generador de Cartas", layout="wide")
    st.title("Generador de Cartas")
    st.caption("Ruteo por YAML, placeholders de texto/imagen, PDF (marca de agua y firma digital opcional), consolidado y validador de calidad.")

    with st.sidebar:
        st.header("Archivos")
        tpl_files = st.file_uploader("Plantillas (.docx)", type=["docx"], accept_multiple_files=True)
        xls_file = st.file_uploader("Base de datos (.xlsx/.xls)", type=["xlsx","xls"])
        img_files = st.file_uploader("Imágenes (logos/firmas)", type=["png","jpg","jpeg"], accept_multiple_files=True)
        st.markdown("---")
        st.header("Opciones")
        newest_first = st.checkbox("Más reciente primero", value=True)
        city = st.text_input("Ciudad (FECHA_CARTA)", value="Medellín")
        letter_date = st.date_input("Fecha a mostrar", value=date.today())
        image_width_in = st.slider("Ancho imágenes (pulgadas)", 0.5, 3.0, 1.5, 0.1)
        st.markdown("---")
        st.header("Routing YAML • Reglas de exportación y derivados")
        yaml_text = st.text_area("Ejemplo:\n"
                                 "templates:\n"
                                 "  - match: 'Hacienda'\n"
                                 "    template: 'MODELO_B.docx'\n"
                                 "    table_index: 1\n"
                                 "    export_pdf: true\n"
                                 "    naming_pattern: 'HAC_{GRUPO}.docx'\n"
                                 "    watermark_text: 'CONFIDENCIAL'\n"
                                 "  - match_regex: '^Secretaría de Gobierno$'\n"
                                 "    template: 'MODELO_A.docx'\n"
                                 "    table_index: 0\n"
                                 "derived_placeholders:\n"
                                 "  SALUDO: '{{PREFIJO}} {{NOMBRE_DIRECTIVO}}'\n"
                                 "footer_text: 'Alcaldía de Medellín — Secretaría General'\n"
                                 "footer_logo_name: 'logo.png'\n", height=280)
        routing_cfg = load_routing_yaml(yaml_text)
        st.markdown("---")
        st.header("Exportación")
        gen_pdf = st.checkbox("Generar PDF", value=False)
        merge_docx = st.checkbox("Consolidar DOCX", value=False)
        merge_pdf = st.checkbox("Consolidar PDF (si se generan PDFs)", value=False)
        add_wm = st.checkbox("Agregar marca de agua (PDF)", value=False)
        wm_text = st.text_input("Texto de marca de agua", value="BORRADOR")
        st.markdown("---")
        st.header("Firma digital (PFX) para PDFs")
        pfx_file = st.file_uploader("Archivo .pfx/.p12", type=["pfx","p12"])
        pfx_pass = st.text_input("Contraseña PFX", type="password")

    if not (tpl_files and xls_file):
        st.info("Sube al menos una plantilla DOCX y el Excel para continuar.")
        return

    # Leer Excel
    try: df = pd.read_excel(xls_file)
    except Exception as e:
        st.error(f"No se pudo leer el Excel: {e}"); return

    st.subheader("Mapeo de columnas")
    auto_map = guess_mapping(df)
    c1, c2, c3 = st.columns(3)
    with c1:
        actor_col = st.selectbox("ACTOR", [None]+list(df.columns), index=(list(df.columns).index(auto_map.get("actor"))+1) if auto_map.get("actor") in df.columns else 0)
        ndir_col  = st.selectbox("NOMBRE DIRECTIVO", [None]+list(df.columns), index=(list(df.columns).index(auto_map.get("nombre_directivo"))+1) if auto_map.get("nombre_directivo") in df.columns else 0)
        pref_col  = st.selectbox("PREFIJO", [None]+list(df.columns), index=(list(df.columns).index(auto_map.get("prefijo"))+1) if auto_map.get("prefijo") in df.columns else 0)
    with c2:
        mesa_col  = st.selectbox("Nombre de la mesa", [None]+list(df.columns), index=(list(df.columns).index(auto_map.get("mesa"))+1) if auto_map.get("mesa") in df.columns else 0)
        nivel_col = st.selectbox("Nivel", [None]+list(df.columns), index=(list(df.columns).index(auto_map.get("nivel"))+1) if auto_map.get("nivel") in df.columns else 0)
        fecha_col = st.selectbox("Fecha", [None]+list(df.columns), index=(list(df.columns).index(auto_map.get("fecha"))+1) if auto_map.get("fecha") in df.columns else 0)
    with c3:
        dato_col  = st.selectbox("Dato transformador", [None]+list(df.columns), index=(list(df.columns).index(auto_map.get("dato"))+1) if auto_map.get("dato") in df.columns else 0)
        firma_col = st.selectbox("Columna archivo firma (opcional)", [None]+list(df.columns), index=(list(df.columns).index(auto_map.get("firma_img"))+1) if auto_map.get("firma_img") in df.columns else 0)
        logo_col  = st.selectbox("Columna archivo logo (opcional)", [None]+list(df.columns), index=(list(df.columns).index(auto_map.get("logo_img"))+1) if auto_map.get("logo_img") in df.columns else 0)

    required = {"actor": actor_col, "nombre_directivo": ndir_col, "prefijo": pref_col,
                "mesa": mesa_col, "nivel": nivel_col, "fecha": fecha_col, "dato": dato_col}
    if any(v is None for v in required.values()):
        st.warning("Completa el mapeo de las columnas requeridas."); return
    if firma_col: required["firma_img"] = firma_col
    if logo_col: required["logo_img"] = logo_col

    # Preparar dataframe
    try:
        work = prepare_dataframe(df, required)
    except Exception as e:
        st.error(str(e)); return

    # ====== Validador de calidad de datos ======
    st.subheader("Calidad de datos")
    colA, colB, colC = st.columns(3)
    with colA:
        st.markdown("**Faltantes por columna (%)**")
        st.dataframe(compute_missing_summary(work), use_container_width=True)
    with colB:
        st.markdown("**Rango de fechas por ACTOR**")
        st.dataframe(compute_date_ranges_by_actor(work), use_container_width=True)
    with colC:
        st.markdown("**Duplicados por ACTOR (si aplica)**")
        dups = compute_duplicates_by_actor(work)
        if dups.empty: st.caption("Sin duplicados.")
        else: st.dataframe(dups, use_container_width=True)

    # Filtro por ACTOR
    st.subheader("Filtro por ACTOR")
    actors = sorted([a for a in work["ACTOR"].dropna().astype(str).unique() if a.strip()])
    sel = st.multiselect("Actores", options=actors, default=actors)
    if sel: work = work[work["ACTOR"].astype(str).isin(sel)]

    with st.expander("Vista previa (ordenada)"):
        work = work.sort_values("_FECHA_TS", ascending=not newest_first, na_position="last")
        st.dataframe(work.drop(columns=["_FECHA_TS"]).head(200), use_container_width=True)

    # Plantillas y assets
    templates_map = {f.name: f.read() for f in tpl_files}
    default_template_bytes = list(templates_map.values())[0]
    image_assets = {f.name: f.read() for f in img_files} if img_files else {}

    # ====== Generación ======
    st.subheader("Generación")
    if st.button("Generar"):
        outputs, errors, index_df = generate_letters_per_group(
            work_df=work,
            default_template_bytes=default_template_bytes,
            templates_map=templates_map,
            routing_cfg=routing_cfg,
            group_field="ACTOR",
            table_index_default=None,
            newest_first=newest_first,
            city=city,
            letter_date=letter_date,
            naming_pattern="CARTA_{GRUPO}.docx",
            image_assets=image_assets,
            image_width_in=image_width_in
        )
        st.success(f"Cartas generadas (DOCX): {len(outputs)}")
        st.dataframe(index_df, use_container_width=True)
        if errors:
            st.warning("Errores:")
            for g, e in errors.items(): st.write(f"- **{g}**: {e}")

        # Descargas DOCX
        for fname, data in outputs.items():
            st.download_button(f"Descargar {fname}", data=data, file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        # Consolidado DOCX
        if merge_docx and outputs:
            from core.merge import merge_documents_docx
            merged = merge_documents_docx(outputs)
            if merged:
                st.download_button("Descargar DOCX consolidado", data=merged, file_name="cartas_consolidado.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        # Reglas de exportación por YAML y switches globales
        pdfs = []
        if gen_pdf or any(r.get("export_pdf") for r in routing_cfg.get("templates", [])):
            for name, docx_b in outputs.items():
                # ¿Regla específica?
                base = name.replace(".docx","")
                # Buscar regla por coincidencia de base en nombre del grupo es complejo; aquí generamos PDF siempre si gen_pdf=True o si routing dice export_pdf:true
                pdf_b = try_docx_to_pdf(docx_b)
                if not pdf_b: 
                    continue
                wm = next((r.get("watermark_text") for r in routing_cfg.get("templates", []) if r.get("export_pdf")), None)
                if add_wm or wm:
                    pdf_b = add_text_watermark(pdf_b, wm or "BORRADOR") or pdf_b
                # Firma digital si se subió PFX
                if pfx_file and pfx_pass:
                    try:
                        pfx_bytes = pfx_file.read()
                        signed = sign_pdf_with_pfx(pdf_b, pfx_bytes, pfx_pass)
                        if signed: pdf_b = signed
                    except Exception:
                        pass
                pdf_name = base + ".pdf"
                pdfs.append((pdf_name, pdf_b))
                st.download_button(f"Descargar {pdf_name}", data=pdf_b, file_name=pdf_name, mime="application/pdf")

        # Consolidado PDF
        if merge_pdf and pdfs:
            merged_pdf = merge_pdfs([b for _, b in pdfs])
            if merged_pdf:
                st.download_button("Descargar PDF consolidado", data=merged_pdf, file_name="cartas_consolidado.pdf", mime="application/pdf")

        # ZIP + Índice
        st.download_button("Descargar todas las cartas (ZIP DOCX)", data=make_zip(outputs),
            file_name="cartas_docx.zip", mime="application/zip")
        st.download_button("Descargar índice (Excel)", data=build_index_sheet(index_df, errors),
            file_name="indice_cartas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
