
# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import Dict, Any, Optional
import re, yaml
from jinja2 import Template

def load_routing_yaml(text: str | None) -> Dict[str, Any]:
    if not text: 
        return {"templates": [], "derived_placeholders": {}}
    try:
        cfg = yaml.safe_load(text) or {}
        if "templates" not in cfg: cfg["templates"] = []
        if "derived_placeholders" not in cfg: cfg["derived_placeholders"] = {}
        return cfg
    except Exception:
        return {"templates": [], "derived_placeholders": {}}

def choose_template_for_group(group_name: str, templates_map: Dict[str, bytes], routing_cfg: Dict[str, Any]) -> tuple[Optional[bytes], Optional[int], Dict[str, Any]]:
    """
    Devuelve (template_bytes, table_index, rule_obj) según reglas.
    Regla puede incluir: template, table_index, export_pdf, naming_pattern, watermark_text, sign_pdf
    """
    if not routing_cfg or "templates" not in routing_cfg: return (None, None, {})
    for rule in routing_cfg["templates"]:
        name = str(group_name)
        hit = False
        if rule.get("match") and rule["match"].lower() in name.lower():
            hit = True
        if not hit and rule.get("match_regex"):
            try:
                if re.search(rule["match_regex"], name):
                    hit = True
            except Exception:
                pass
        if hit:
            tpl_name = rule.get("template")
            idx = rule.get("table_index")
            tpl_bytes = templates_map.get(tpl_name) if tpl_name in templates_map else None
            return (tpl_bytes, idx, rule)
    return (None, None, {})

def render_derived_placeholders(mapping: Dict[str,str], derived_cfg: Dict[str,str]) -> Dict[str,str]:
    """
    derived_cfg ejemplo:
      SALUDO: "{{PREFIJO}} {{NOMBRE_DIRECTIVO}}"
      DESTINATARIO: "{{NOMBRE_DIRECTIVO}}"
    """
    out = {}
    for k, expr in (derived_cfg or {}).items():
        try:
            tmpl = Template(str(expr))
            out[k] = tmpl.render(**mapping)
        except Exception:
            out[k] = ""
    return out
