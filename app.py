import io
import re
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel-Duplikate prüfen & bereinigen", layout="wide")
st.title("Excel-Duplikate prüfen & bereinigen")

st.write("""
**Prüfung A:** Duplikate, wenn **alle Spalten außer `externalId`** identisch sind  
⇒ Downloads: bereinigte Datei + entfernte `externalId`s

**Prüfung B:** Duplikate, wenn die **Kernfelder** identisch sind, wobei `name` & Co. **normalisiert** werden:  
`taxExemption, name, recipientCountry, category, vendorCountry, itemsTaxRate`  
(also **ohne** `externalId` und **ohne** `code`).  
⇒ Download: **Liste der `externalId`s** aller Zeilen, die zu solchen Duplikat-Gruppen gehören.
""")

# ------------------------------------------------
# Helpers: Einlesen & Export
# ------------------------------------------------

def read_excel_as_str(file) -> pd.DataFrame:
    """Excel einlesen (alle Spalten als str), führende Nullen bleiben erhalten."""
    df = pd.read_excel(file, dtype=str)
    df = df.fillna("")
    # trim
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()

def highlight_mask(df: pd.DataFrame, mask):
    styles = pd.DataFrame("", index=df.index, columns=df.columns)
    styles.loc[mask, :] = "background-color: #ffd6cc"  # hellrot/orange
    return styles

# ------------------------------------------------
# Prüfung A: alle Spalten außer externalId
# ------------------------------------------------

def run_check_A(df: pd.DataFrame, id_col="externalId"):
    if id_col not in df.columns:
        raise ValueError(f"Spalte '{id_col}' fehlt.")
    compare_cols = [c for c in df.columns if c != id_col]

    dup_mask = df.duplicated(subset=compare_cols, keep=False)

    # stabil sortieren: Keys + externalId (string, behält führende Nullen)
    df_sorted = df.copy()
    df_sorted["_idx"] = range(len(df_sorted))
    df_sorted = df_sorted.sort_values(by=[*compare_cols, id_col], kind="mergesort")
    df_sorted["_keep"] = ~df_sorted.duplicated(subset=compare_cols, keep="first")

    cleaned_df = df_sorted[df_sorted["_keep"]].drop(columns=["_idx", "_keep"]).reset_index(drop=True)
    removed_rows_sorted = df_sorted[~df_sorted["_keep"]]
    removed_ids_df = removed_rows_sorted[[id_col]].drop_duplicates().reset_index(drop=True)

    removed_count = len(removed_rows_sorted)
    group_count = df.loc[dup_mask, compare_cols].drop_duplicates().shape[0] if dup_mask.any() else 0
    return dup_mask, cleaned_df, removed_ids_df, removed_count, group_count

# ------------------------------------------------
# Prüfung B: Kernfelder mit Normalisierung
# ------------------------------------------------

CORE_FIELDS = [
    "taxExemption",
    "name",
    "recipientCountry",
    "category",
    "vendorCountry",
    "itemsTaxRate",
]

# -- Normalisierer

_REM_WORDS = {
    "services", "service", "fa/ooe", "fa\\ooe", "fa ooe"
}

def _to_ascii_lower(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def _norm_bool(s: str) -> str:
    s_l = s.strip().lower()
    return "true" if s_l in {"true", "1", "yes", "y", "wahr"} else "false" if s_l in {"false", "0", "no", "n", "falsch"} else s_l

def _norm_rate(s: str) -> str:
    # 8,1% / 0.081 → numerisch robust
    t = s.strip().replace("%", "")
    t = t.replace(",", ".")
    try:
        val = float(t)
        # falls > 1 als Prozent interpretieren (8.1 -> 0.081)
        if val > 1.0:
            val = val / 100.0
        return f"{val:.6f}"
    except Exception:
        return _to_ascii_lower(s).strip()

def _norm_name(s: str) -> str:
    if not isinstance(s, str):
        s = str(s)
    s = _to_ascii_lower(s)

    # Klammerinhalte entfernen
    s = re.sub(r"\([^)]*\)", " ", s)

    # lange/kurze Gedankenstriche & Bindestriche in Spaces umwandeln
    s = s.replace("–", " ").replace("—", " ").replace("-", " ")

    # nicht-alphanumerische Zeichen (außer Leerzeichen & Komma/Punkt für Zahlen) entfernen
    s = re.sub(r"[^a-z0-9\s\.,]", " ", s)

    # Stoppwörter entfernen (services/service/fa/ooe)
    for w in _REM_WORDS:
        s = re.sub(rf"\b{re.escape(w)}\b", " ", s)

    # Zahlenformate vereinheitlichen (8,1 -> 8.1)
    s = s.replace(",", ".")

    # Mehrfach-Whitespace zusammenfassen & trimmen
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _norm_text(s: str) -> str:
    # für Länder/Category etc.: ascii+lower+trim
    return re.sub(r"\s+", " ", _to_ascii_lower(s)).strip()

def normalize_core_view(df: pd.DataFrame) -> pd.DataFrame:
    """Erzeugt eine normalisierte Sicht nur der CORE_FIELDS."""
    missing = [c for c in CORE_FIELDS if c not in df.columns]
    if missing:
        raise ValueError(f"Folgende Spalten fehlen für Prüfung B: {', '.join(missing)}")

    view = pd.DataFrame(index=df.index)
    view["taxExemption"] = df["taxExemption"].map(_norm_bool)
    view["name"] = df["name"].map(_norm_name)
    view["recipientCountry"] = df["recipientCountry"].map(_norm_text)
    view["category"] = df["category"].map(_norm_text)
    view["vendorCountry"] = df["vendorCountry"].map(_norm_text)
    view["itemsTaxRate"] = df["itemsTaxRate"].map(_norm_rate)
    return view

def run_check_B(df: pd.DataFrame, id_col="externalId"):
    norm = normalize_core_view(df)

    # Duplikate nach normalisierten Kernfeldern
    dup_mask_core = norm.duplicated(subset=CORE_FIELDS, keep=False)

    # externalIds der Zeilen, die in Duplikat-Gruppen fallen
    external_ids_in_groups = df.loc[dup_mask_core, id_col].drop_duplicates().reset_index(drop=True)

    # Ansicht der Gruppen (normalisierte Schlüssel + original externalId)
    view = norm.copy()
    view[id_col] = df[id_col]
    duplicates_view = view.loc[dup_mask_core, CORE_FIELDS + [id_col]] \
                          .sort_values(CORE_FIELDS + [id_col], kind="mergesort")

    return dup_mask_core, external_ids_in_groups, duplicates_view

# ------------------------------------------------
# UI
# ------------------------------------------------

uploaded = st.file_uploader("Excel-Datei hochladen (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        df_input = read_excel_as_str(uploaded)
    except Exception as e:
        st.error(f"Fehler beim Laden der Datei: {e}")
        st.stop()

    st.subheader("Vorschau")
    st.dataframe(df_input.head(50), use_container_width=True)

    if st.button("Prüfungen ausführen"):
        try:
            # ---------- Prüfung A ----------
            st.markdown("## Prüfung A – Duplikate (alle Spalten außer `externalId`)")

            dup_mask_A, cleaned_df, removed_ids_df, removed_count, group_count = run_check_A(df_input, id_col="externalId")
            st.success(f"Prüfung A: {removed_count} Zeile(n) entfernt in {group_count} Duplikat-Gruppe(n).")

            if dup_mask_A.any():
                st.subheader("Gefundene Duplikate (A) – farblich markiert")
                st.dataframe(
                    df_input.style.apply(lambda _: highlight_mask(df_input, dup_mask_A), axis=None),
                    use_container_width=True,
                )
            else:
                st.info("Prüfung A: Keine Duplikate gefunden.")

            st.subheader("Bereinigte Tabelle (A) – Export-Vorschau")
            st.dataframe(cleaned_df, use_container_width=True)

            # Downloads A
            cleaned_bytes = to_excel_bytes(cleaned_df, sheet_name="Bereinigt")
            st.download_button(
                label="(A) Bereinigte Excel herunterladen (.xlsx)",
                data=cleaned_bytes,
                file_name="bereinigt_ohne_duplikate.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            removed_bytes = to_excel_bytes(removed_ids_df, sheet_name="Entfernte_externalIds")
            st.download_button(
                label="(A) Entfernte externalIds herunterladen (.xlsx)",
                data=removed_bytes,
                file_name="entfernte_externalIds.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.markdown("---")

            # ---------- Prüfung B ----------
            st.markdown("## Prüfung B – Duplikate nach normalisierten Kernfeldern (ohne `externalId` & `code`)")

            dup_mask_B, ext_ids_core_dups, duplicates_view_B = run_check_B(df_input, id_col="externalId")

            if dup_mask_B.any():
                st.success(f"Prüfung B: {len(ext_ids_core_dups)} `externalId`(s) gehören zu Duplikat-Gruppen basierend auf normalisierten Kernfeldern.")
                st.subheader("Ansicht der Duplikat-Gruppen (B) – normalisierte Schlüssel + externalId")
                st.dataframe(duplicates_view_B, use_container_width=True)

                # Download nur externalIds (B)
                ext_ids_B_bytes = to_excel_bytes(ext_ids_core_dups.to_frame(name="externalId"), sheet_name="externalIds_Duplikate_B")
                st.download_button(
                    label="(B) externalIds der Duplikate (Kernfelder) herunterladen (.xlsx)",
                    data=ext_ids_B_bytes,
                    file_name="externalIds_duplikate_kernfelder.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("Prüfung B: Keine Duplikate nach den Kernfeldern gefunden.")

        except Exception as e:
            st.error(f"Fehler bei der Verarbeitung: {e}")

else:
    st.info("Bitte eine Excel-Datei (.xlsx) hochladen, um zu starten.")
