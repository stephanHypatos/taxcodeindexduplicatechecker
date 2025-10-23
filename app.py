import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel-Duplikate prüfen & bereinigen", layout="wide")

st.title("Excel-Duplikate prüfen & bereinigen")
st.write("""
Lade eine **Excel-Datei (.xlsx)** hoch. Duplikate sind Zeilen, bei denen **alle Spalten außer `externalId`** identisch sind.  
Führende Nullen in `externalId` bleiben erhalten (wird als String gelesen).

**Downloads am Ende:**
1) Bereinigte Excel ohne Duplikate  
2) Datei mit den **entfernten `externalId`s**
""")

# -------------------------------
# Helpers
# -------------------------------

def read_excel_as_str(file) -> pd.DataFrame:
    """Liest Excel ein, behält führende Nullen (dtype=str) in allen Spalten."""
    df = pd.read_excel(file, dtype=str)
    df = df.fillna("")
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def highlight_duplicates(df: pd.DataFrame, dup_mask):
    styles = pd.DataFrame("", index=df.index, columns=df.columns)
    styles.loc[dup_mask, :] = "background-color: #ffd6cc"  # hellrot/orange
    return styles

def to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()

def find_duplicates_and_outputs(df: pd.DataFrame, id_col="externalId"):
    if id_col not in df.columns:
        raise ValueError(f"Spalte '{id_col}' fehlt.")

    compare_cols = [c for c in df.columns if c != id_col]

    # Duplikate nach Vergleichsspalten
    dup_mask = df.duplicated(subset=compare_cols, keep=False)

    # Stabil sortieren: erst nach Vergleichsschlüsseln, dann externalId (als String, behält führende Nullen)
    df_sorted = df.copy()
    df_sorted["_orig_idx"] = range(len(df_sorted))
    df_sorted = df_sorted.sort_values(by=[*compare_cols, id_col], kind="mergesort")

    # In jeder Gruppe (compare_cols) die erste Zeile behalten
    df_sorted["_keep"] = ~df_sorted.duplicated(subset=compare_cols, keep="first")

    # Bereinigte Tabelle = nur _keep == True
    cleaned_sorted = df_sorted[df_sorted["_keep"]].drop(columns=["_orig_idx", "_keep"])
    cleaned_df = cleaned_sorted.reset_index(drop=True)

    # Entfernte Zeilen = _keep == False
    removed_rows_sorted = df_sorted[~df_sorted["_keep"]]
    removed_ids_df = (
        removed_rows_sorted[[id_col]]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    removed_count = len(removed_rows_sorted)
    group_count = (
        df.loc[dup_mask, compare_cols].drop_duplicates().shape[0]
        if dup_mask.any() else 0
    )

    return dup_mask, cleaned_df, removed_ids_df, removed_count, group_count


# -------------------------------
# UI
# -------------------------------

uploaded = st.file_uploader("Excel-Datei hochladen (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        df_input = read_excel_as_str(uploaded)
    except Exception as e:
        st.error(f"Fehler beim Laden der Datei: {e}")
        st.stop()

    st.subheader("Vorschau")
    st.dataframe(df_input.head(50), use_container_width=True)

    if st.button("Duplikate prüfen & bereinigen"):
        try:
            dup_mask, cleaned_df, removed_ids_df, removed_count, group_count = find_duplicates_and_outputs(df_input)

            st.success(f"{removed_count} Zeile(n) entfernt in {group_count} Duplikat-Gruppe(n).")

            if not dup_mask.any():
                st.info("Keine Duplikate gefunden (nach Regel: alle Spalten außer 'externalId' identisch).")
            else:
                st.subheader("Gefundene Duplikate (farblich markiert)")
                st.dataframe(
                    df_input.style.apply(lambda _: highlight_duplicates(df_input, dup_mask), axis=None),
                    use_container_width=True,
                )

            st.subheader("Bereinigte Tabelle (Export-Vorschau)")
            st.dataframe(cleaned_df, use_container_width=True)

            # --- Downloads ---
            cleaned_bytes = to_excel_bytes(cleaned_df, sheet_name="Bereinigt")
            st.download_button(
                label="Bereinigte Excel herunterladen (.xlsx)",
                data=cleaned_bytes,
                file_name="bereinigt_ohne_duplikate.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            removed_bytes = to_excel_bytes(removed_ids_df, sheet_name="Entfernte_externalIds")
            st.download_button(
                label="Entfernte externalIds herunterladen (.xlsx)",
                data=removed_bytes,
                file_name="entfernte_externalIds.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Fehler bei der Verarbeitung: {e}")
else:
    st.info("Bitte eine Excel-Datei (.xlsx) hochladen, um zu starten.")
