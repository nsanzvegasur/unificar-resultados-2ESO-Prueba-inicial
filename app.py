import io
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Unificar resultados · 2.º ESO", page_icon="📊", layout="wide")

st.title("Unificar resultados · Prueba inicial 2.º ESO")
st.write("Sube los Excel de toda la clase y la aplicación los reunirá en un único archivo. Después podrás introducir manualmente la nota de producción escrita y descargar el resultado completo.")

AREAS = ["Comprensión", "Morfología", "Semántica", "Textos", "Literatura", "Sintaxis"]
AREA_KEYS = {
    "Comprensión": "comprension", "Morfología": "morfologia", "Semántica": "semantica",
    "Textos": "textos", "Literatura": "literatura", "Sintaxis": "sintaxis"
}


def norm(s):
    s = str(s).strip().lower()
    s = s.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").replace("ñ", "n")
    return re.sub(r"[^a-z0-9]+", "", s)


def number(v):
    if pd.isna(v) or str(v).strip() == "":
        return None
    s = str(v).strip().replace("%", "").replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


def find_col(df, candidates):
    wanted = {norm(x) for x in candidates}
    for c in df.columns:
        if norm(c) in wanted:
            return c
    return None


def extract_from_table(df, filename):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    out = []
    name_col = find_col(df, ["name", "nombre", "alumno", "nombre y apellidos", "nombreyapellidos"])
    group_col = find_col(df, ["group", "grupo"])
    date_col = find_col(df, ["date", "fecha", "fecha y hora", "fechayhora"])
    if name_col:
        for _, r in df.iterrows():
            name = str(r.get(name_col, "")).strip()
            if not name or name.lower() in {"nan", "none"}:
                continue
            item = {
                "Alumno": name,
                "Grupo": r.get(group_col, "") if group_col else "",
                "Fecha": r.get(date_col, "") if date_col else "",
                "Fuente": filename,
            }
            for area, key in AREA_KEYS.items():
                col = find_col(df, [key, area])
                item[area] = number(r.get(col)) if col else None
            for label, candidates in {
                "Nota sobre 9": ["nota_final_sobre_9", "nota de esta parte (sobre 9)", "nota sobre 9"],
                "Ortografía": ["faltas_ortografia", "faltas de ortografía detectadas"],
                "Tildes": ["faltas_tildes", "faltas de tilde detectadas"],
            }.items():
                col = find_col(df, candidates)
                item[label] = number(r.get(col)) if col else None
            out.append(item)
    return out


def extract_vertical(df, filename):
    rows = []
    for _, r in df.iterrows():
        vals = [str(x).strip() if not pd.isna(x) else "" for x in r.tolist()]
        if vals:
            rows.append(vals)
    data = {}
    for vals in rows:
        if len(vals) >= 2:
            data[norm(vals[0])] = vals[1]

    name = data.get("alumno") or data.get("nombre") or data.get("nombreyapellidos")
    if not name:
        return []

    item = {
        "Alumno": name.strip(),
        "Grupo": data.get("grupo", ""),
        "Fecha": data.get("fechayhora", data.get("fecha", "")),
        "Fuente": filename,
    }
    for area, key in AREA_KEYS.items():
        val = data.get(key, data.get(norm(area)))
        item[area] = number(val)

    for label, keys in {
        "Nota sobre 9": ["notadeestapartesobre9", "notasobre9", "notafinalsobre9"],
        "Ortografía": ["faltasdeortografiadetectadas", "faltasortografia"],
        "Tildes": ["faltasdetildedetectadas", "faltastildes"],
    }.items():
        item[label] = number(next((data[k] for k in keys if k in data), None))
    return [item]


def read_excel(upload):
    sheets = pd.read_excel(upload, sheet_name=None, header=0)
    records = []
    for _, df in sheets.items():
        table_records = extract_from_table(df, upload.name)
        if table_records:
            records.extend(table_records)
        else:
            records.extend(extract_vertical(df, upload.name))
    return records


uploads = st.file_uploader("Sube los Excel de los alumnos", type=["xlsx", "xls"], accept_multiple_files=True)

if uploads:
    all_records = []
    errors = []
    for upload in uploads:
        try:
            all_records.extend(read_excel(upload))
        except Exception as e:
            errors.append(f"{upload.name}: {e}")

    if errors:
        st.warning("Algunos archivos no se han podido leer: " + " | ".join(errors))

    if all_records:
        df = pd.DataFrame(all_records)
        for c in AREAS + ["Nota sobre 9", "Ortografía", "Tildes"]:
            if c not in df.columns:
                df[c] = None

        df["Producción escrita"] = None
        df["Nota final sobre 10"] = None
        df = df[["Alumno", "Grupo", "Fecha"] + AREAS + ["Nota sobre 9", "Producción escrita", "Nota final sobre 10", "Ortografía", "Tildes", "Fuente"]]

        st.success(f"Se han encontrado {len(df)} resultados.")
        st.subheader("Notas de producción escrita")
        st.caption("Introduce manualmente la nota de producción escrita de cada alumno, de 0 a 1. La nota final será la suma de Nota sobre 9 + Producción escrita.")

        edit_cols = ["Alumno", "Grupo"] + AREAS + ["Nota sobre 9", "Producción escrita"]
        edited = st.data_editor(
            df[edit_cols],
            hide_index=True,
            use_container_width=True,
            column_config={
                "Producción escrita": st.column_config.NumberColumn(
                    "Producción escrita (0–1)", min_value=0.0, max_value=1.0, step=0.05, format="%.2f"
                ),
            },
        )

        # La nota final se calcula automáticamente como Nota sobre 9 + Producción escrita.
        edited["Nota final sobre 10"] = (
            edited["Nota sobre 9"].fillna(0) + edited["Producción escrita"].fillna(0)
        ).clip(upper=10).round(2)

        st.subheader("Resultado final")
        result_cols = ["Alumno", "Grupo", "Nota sobre 9", "Producción escrita", "Nota final sobre 10"]
        st.dataframe(edited[result_cols], hide_index=True, use_container_width=True)

        st.subheader("Media de la clase por apartados")
        chart_df = edited[AREAS].mean().rename("Media").to_frame()
        chart_df.loc["Producción escrita"] = edited["Producción escrita"].mean() * 10
        st.bar_chart(chart_df, y="Media")
        st.caption("El gráfico incluye Comprensión, Morfología, Semántica, Textos, Literatura, Sintaxis y Producción escrita. La producción se muestra sobre 10 para hacerla comparable visualmente con las demás áreas.")

        final_df = df.copy()
        final_df["Producción escrita"] = edited["Producción escrita"]
        final_df["Nota final sobre 10"] = edited["Nota final sobre 10"]
        final_df = final_df.drop(columns=["Fuente"])

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            final_df.to_excel(writer, sheet_name="Resultados", index=False)
            chart_df.round(2).to_excel(writer, sheet_name="Medias por áreas")

        st.download_button(
            "DESCARGAR EXCEL DE LA CLASE",
            buffer.getvalue(),
            file_name="Resultados_unificados_2ESO_Prueba_inicial.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.error("No se ha encontrado ningún resultado reconocible en los archivos. Comprueba que sean los Excel de los resultados individuales.")
else:
    st.info("Sube los Excel individuales para empezar.")
