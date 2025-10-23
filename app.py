import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel-Duplikate prüfen & bereinigen", layout="wide")

st.title("Excel-Duplikate prüfen & bereinigen")
st.write(
    """
Lade eine **Excel-Datei (.xlsx)** mit den Spalten (z. B.)  
`externalId, code, taxExemption, name, recipientCountry, category, vendorCountry, itemsTaxRate, Column1` hoch.

**Logik:** Zeilen gelten als Duplikate, wenn **alle Spalten außer `externalId` identisch** sind.  
Beim Bereinigen bleibt **nur die Zeile mit der kleinsten `externalId`** pro Duplikat-Gruppe erhalten.
"""
)

# ---------- Helper ----------
def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Trimmt Strings, normalisiert leere Strings auf NaN, lässt andere Typen unverändert."""
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            # strip whitespaces
            df[col] = df[col].astype(str).str.strip()
            # convert empty strings to NA
            df[col] = df[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    return df

def to_numeric_safe(s: pd.Series):
    """Konvertiert externalId bestmöglich numerisch, um korrekt die kleinste ID zu bestimmen.
    Fällt zurück auf String-Vergleich, wenn nicht numerisch."""
    num = pd.to_numeric(s, errors="coerce")
    if num.notna().any():
        # Fülle NaN mit +inf, damit echte Zahlen bevorzugt klein sind
        return num.fillna(float("inf"))
    # rein lexikografisch vergleichen, falls alles nicht-numerisch
    return s.astype(str)

def find_duplicates_and_clean(df: pd.DataFrame, id_col: str = "externalId"):
    if id_col not in df.columns:
        raise ValueError(f"Spalte '{id_col}' fehlt in der Datei.")

    # Normalisieren (optional, aber sehr nützlich)
    df_norm = normalize_dataframe(df)

    # Spalten, nach denen Duplikate definiert werden (alle außer externalId)
    key_cols = [c for c in df_norm.columns if c != id_col]
    if not key_cols:
        raise ValueError("Es gibt keine Vergleichsspalten außer 'externalId'.")

    # Maske: alle Zeilen, die zu einer Duplikat-Gruppe gehören (>=2 gleich)
    dup_mask = df_norm.duplicated(subset=key_cols, keep=False)

    duplicates_df = df_norm.loc[dup_mask].copy()

    # Bereinigung: je Gruppe die kleinste externalId (numerisch bevorzugt) behalten
    df_for_sort = df_norm.copy()
    df_for_sort["_externalId_sort"] = to_numeric_safe(df_for_sort[id_col])

    # Drop duplicates by key_cols, keeping the first after sorting by keys + _externalId_sort + externalId
    df_for_sort = df_for_sort.sort_values(key_cols + ["_externalId_sort", id_col], kind="mergesort")
    cleaned_df = df_for_sort.drop_duplicates(subset=key_cols, keep="first").drop(columns=["_externalId_sort"])

    # Statistiken
    removed_count = len(df_norm) - len(cleaned_df)
    group_count = (
        duplicates_df[key_cols]
        .drop_duplicates()
        .shape[0]
        if not duplicates_df.empty
        else 0
    )

    return duplicates_df, cleaned_df, removed_count, group_count

def df_to_excel_download(df: pd.DataFrame, sheet_name: str = "bereinigt") -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()

# ---------- UI ----------
uploaded = st.file_uploader("Excel-Datei hochladen (.xlsx)", type=["xlsx"])

with st.expander("Optionen"):
    keep_original_order = st.checkbox(
        "Original-Reihenfolge im Export erhalten (statt nach Gruppen/IDs sortiert)",
        value=False,
    )

if uploaded:
    try:
        df_input = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Fehler beim Laden der Datei: {e}")
        st.stop()

    st.subheader("Vorschau")
    st.dataframe(df_input.head(50), use_container_width=True)

    if st.button("Duplikate prüfen & bereinigen"):
        try:
            duplicates_df, cleaned_df, removed_count, group_count = find_duplicates_and_clean(df_input, id_col="externalId")

            st.success(
                f"Ergebnis: {removed_count} Zeile(n) entfernt in {group_count} Duplikat-Gruppe(n)."
            )

            if duplicates_df.empty:
                st.info("Keine Duplikate gefunden (nach Regel: alle Spalten außer 'externalId' identisch).")
            else:
                st.subheader("Gefundene Duplikate")
                # Zur leichteren Sichtbarkeit sortieren wir die Duplikate nach den Key-Spalten + externalId
                key_cols = [c for c in duplicates_df.columns if c != "externalId"]
                show_dups = duplicates_df.sort_values(key_cols + ["externalId"], kind="mergesort")
                st.dataframe(show_dups, use_container_width=True)

            # Export vorbereiten
            if keep_original_order:
                # Reihenfolge wie im Input beibehalten -> Reindex nach Originalindex
                cleaned_df = df_input.merge(
                    cleaned_df.assign(_keep=1),
                    how="inner",
                    on=list(cleaned_df.columns),
                ).drop(columns=["_keep"])
                # Falls obiger Merge wegen identischer Spalten schwierig: alternativ per Key-Vergleichsbasis neu sortieren
            else:
                # Sinnvolle Default-Sortierung: nach Vergleichsschlüsseln + externalId
                sort_cols = [c for c in cleaned_df.columns if c != "externalId"] + ["externalId"]
                cleaned_df = cleaned_df.sort_values(sort_cols, kind="mergesort")

            st.subheader("Bereinigte Tabelle (Export-Vorschau)")
            st.dataframe(cleaned_df, use_container_width=True)

            xlsx_bytes = df_to_excel_download(cleaned_df, sheet_name="bereinigt")
            st.download_button(
                label="Bereinigte Excel herunterladen (.xlsx)",
                data=xlsx_bytes,
                file_name="bereinigt_ohne_duplikate.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Fehler bei der Verarbeitung: {e}")

else:
    st.info("Bitte eine .xlsx-Datei hochladen, um zu starten.")
