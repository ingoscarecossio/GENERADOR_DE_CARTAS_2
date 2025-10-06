
from .backend import (
    guess_mapping, prepare_dataframe, parse_date, format_date_dmy, slugify,
    list_candidate_tables, find_target_table, clear_table_keep_header, fill_table, month_name_es
)
from .routing import choose_template_for_group, load_routing_yaml, render_derived_placeholders
from .funcionalidades import (
    generate_letters_per_group, build_index_sheet, make_zip
)
from .pdf_utils import try_docx_to_pdf, merge_pdfs, add_text_watermark, sign_pdf_with_pfx
from .merge import merge_documents_docx
from .quality import compute_missing_summary, compute_duplicates_by_actor, compute_date_ranges_by_actor

__all__ = [
    "guess_mapping","prepare_dataframe","parse_date","format_date_dmy","slugify",
    "list_candidate_tables","find_target_table","clear_table_keep_header","fill_table","month_name_es",
    "choose_template_for_group","load_routing_yaml","render_derived_placeholders",
    "generate_letters_per_group","build_index_sheet","make_zip",
    "try_docx_to_pdf","merge_pdfs","add_text_watermark","sign_pdf_with_pfx",
    "merge_documents_docx",
    "compute_missing_summary","compute_duplicates_by_actor","compute_date_ranges_by_actor"
]
