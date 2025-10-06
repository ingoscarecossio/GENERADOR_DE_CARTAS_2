
# -*- coding: utf-8 -*-
from __future__ import annotations
import pandas as pd
import numpy as np

def compute_missing_summary(df: pd.DataFrame) -> pd.DataFrame:
    miss = df.isna().mean().rename("Porcentaje faltantes").to_frame()
    miss["Porcentaje faltantes"] = (miss["Porcentaje faltantes"] * 100).round(2)
    return miss.sort_values("Porcentaje faltantes", ascending=False)

def compute_duplicates_by_actor(df: pd.DataFrame) -> pd.DataFrame:
    if "ACTOR" not in df.columns: return pd.DataFrame()
    cols = [c for c in ["ACTOR","MESA","NIVEL","FECHA","DATO"] if c in df.columns]
    dups = df[df.duplicated(subset=cols, keep=False)].sort_values(["ACTOR","MESA","FECHA"])
    return dups[cols]

def compute_date_ranges_by_actor(df: pd.DataFrame) -> pd.DataFrame:
    if "ACTOR" not in df.columns or "_FECHA_TS" not in df.columns: return pd.DataFrame()
    ag = df.groupby("ACTOR")["_FECHA_TS"].agg(["min","max","count"]).reset_index()
    ag.columns = ["ACTOR","Fecha mínima","Fecha máxima","Registros"]
    return ag.sort_values("ACTOR")
