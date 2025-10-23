import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel-Duplikate prüfen & bereinigen", layout="wide")

st.title("Excel-Duplikate prüfen & bereinigen")

st.write("""
Lade eine **Excel-Datei (.xlsx)** mit Spalten wie  
`externalId, code, taxExemption, name, recipientCountry, category, vendorCountry, itemsTaxRate, Column1`.

**Logik:** Zeilen gelten als Duplikate, wenn **alle Spalten außer `externalId` identisch** sind.  
Die App markiert diese Zeilen und exportiert eine bereinigte Datei, in der **nur eine Zeile pro Gruppe** bleibt
(standardmäßig die mit der kleinsten `externalId`, wobei führende Nullen beibehalten werden).
""")

# -------------------------------
# Helper-Funktionen
# -------------------------------

def read_excel_as_str(file) -> pd.DataFrame:
    """Liest Excel ein, behält führende Nullen in externalId."""
    df = pd.read_excel(file, dtype=str)
    df = df.fillna("")  # Leere Zellen zu leeren Strings
    # Alle Spalten trimmen
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def find_duplicates_and_clean(df: pd.DataFrame, id_col="externalId"):
    if id_col not in df.columns:
        raise ValueError(f"Spalte '{id_col}' fehlt in der Datei.")

    # Spalten für Vergleich
    compare_cols = [c for c in df.columns if c != id_col]

    # Duplikate finden (nach allen anderen Spalten)
    dup_mask = df.duplicated(subset=compare_cols, keep=False)

    # Nur Duplikate anzeigen
    duplicates_df = df.loc[dup_mask].copy()

    # Für jede Gruppe nur kleinste externalId (stringvergleich, behält führende 0)
    cleaned_df = (
        df.sort_values(by=[*compare_cols, id_col], kind="mergesort")
        .drop_duplicates(subset=compare_cols, keep="first")
        .reset_index(drop=True)
    )

    removed_count = len(df) - len(cleaned_df)
    group_count = (
        duplicates_df[compare_cols].drop_duplicates().shape[0]
        if not duplicates_df.empty
        else 0
    )

    return duplicates_df, cleaned_df, removed_count, group_count, dup_mask

def highlight_duplicates(df: pd.DataFrame, dup_mask):
    """Markiert Duplikate farblich (orange)."""
    styles = pd.DataFrame("", index=df.index, columns=df.columns)
    styles.loc[dup_mask, :] = "background-color: #ffd6cc"  # hellrot
    return styles

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Exportiert DataFrame zu Excel Bytes."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Bereinigt")
    return buffer.getvalue()


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
            duplicates_df, cleaned_df, removed_count, group_count, dup_mask = find_duplicates_and_clean(df_input)

            st.success(
                f"{removed_count} Zeile(n) entfernt in {group_count} Duplikat-Gruppe(n)."
            )

            if duplicates_df.empty:
                st.info("Keine Duplikate gefunden.")
            else:
                st.subheader("Gefundene Duplikate (farblich markiert)")
                st.dataframe(
                    df_input.style.apply(lambda _: highlight_duplicates(df_input, dup_mask), axis=None),
                    use_container_width=True,
                )

            st.subheader("Bereinigte Tabelle (Export-Vorschau)")
            st.dataframe(cleaned_df, use_container_width=True)

            excel_data = df_to_excel_bytes(cleaned_df)
            st.download_button(
                label="Bereinigte Excel herunterladen (.xlsx)",
                data=excel_data,
                file_name="bereinigt_ohne_duplikate.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Fehler bei der Verarbeitung: {e}")
else:
    st.info("Bitte eine Excel-Datei (.xlsx) hochladen, um zu starten.")
