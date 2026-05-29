import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Control Aforos BYMA")
    st.markdown(
        "Compara los aforos informados por la API BYMA (col 26 de ESPECIES.XLS) "
        "con las listas asignadas en Gallo. Detecta diferencias de lista/aforo, "
        "especies aceptadas por BYMA sin lista en Gallo, y genera un **Reporte Garantías BYMA** "
        "con todas las especies aceptadas agrupadas por categoría y buscador por ticker/CVSA."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Archivo de referencia** (se toma automáticamente de Archivos Compartidos):
        - **ESPECIES.XLS** — maestro de especies: código CVSA, Lista asignada (col F)
          y haircut BYMA API (col 26). Aforo BYMA = 100 − haircut.

        **Output 1 — `DIFERENCIAS AFOROS.xlsx`:**
        - Hoja **Diferencias Aforo**: especies cuya Lista en Gallo no coincide con el aforo BYMA.
          Columna *Aforo Lista Actual* en rojo, *Lista Sugerida* en verde (corrección).
        - Hoja **Sin Lista Gallo**: especies que BYMA acepta pero no tienen lista en Gallo.
        - Hoja **Lista Sin BYMA** *(informacional)*: especies con lista en Gallo que BYMA ya no acepta.

        **Output 2 — `Reporte Garantias BYMA DD-MM-AA.xlsx`:**
        - Hoja **Resumen**: tabla de totales por categoría (cantidad, aforo mín/máx,
          disponibilidad ARS / USD MEP / USD Cable) + buscador rápido por ticker o código CVSA.
        - Una hoja por categoría: Títulos Públicos, Letras del Tesoro, Obligaciones Negociables,
          Acciones, CEDEARs, FCI, Otros. Incluye ticker ARS / MEP / Cable, vencimiento/emisor,
          haircut y aforo. Colores según nivel de aforo.

        **Tablas Lista → Aforo:**

        | Segmento | Listas | Aforos |
        |---|---|---|
        | Renta Variable | 1–8 | 85 %–30 % |
        | Renta Fija Públicos | 10–17 | 95 %–60 % |
        | Renta Fija Privados | 22–27 | 85 %–60 % |
        | Letras y Bonos del Tesoro | 85 / 90 / 95 | 85 % / 90 % / 95 % |
        """)

    # ── Input (desde Archivos Compartidos) ────────────────────────────────────
    especies_file = shared_or_upload(
        "shared_especies", "Maestro de especies (ESPECIES.XLS)",
        ["xls", "xlsx"], "caf_especies"
    )

    # ── Session state ─────────────────────────────────────────────────────────
    if "caf_result" not in st.session_state:
        st.session_state["caf_result"] = None

    st.divider()

    if not especies_file:
        st.info("Cargá ESPECIES.XLS desde Archivos Compartidos para ejecutar el control.")
        return

    # ── Ejecución ─────────────────────────────────────────────────────────────
    if st.button("Ejecutar control", type="primary", use_container_width=True):
        with st.spinner("Procesando maestro de especies..."):
            try:
                from skills.control_aforos_byma.logic import generar_control
                xlsx_dif_bytes, xlsx_rc_bytes, resumen, advertencias = generar_control(especies_file)
                st.session_state["caf_result"] = {
                    "xlsx_dif": xlsx_dif_bytes,
                    "xlsx_rc":  xlsx_rc_bytes,
                    "resumen":  resumen,
                    "advertencias": advertencias,
                }
            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())

    # ── Resultados — persisten en session_state entre reruns ──────────────────
    result = st.session_state.get("caf_result")
    if not result:
        return

    resumen      = result["resumen"]
    advertencias = result["advertencias"]
    n_dif = resumen["diferencias"]
    n_sl  = resumen["sin_lista"]
    n_lsb = resumen["lista_sin_byma"]

    if n_dif == 0 and n_sl == 0:
        st.success(
            f"✓ Sin diferencias ni faltantes — maestro alineado con BYMA API.  "
            f"({resumen['total_byma']} especies en lista BYMA)"
        )
    else:
        partes = []
        if n_dif: partes.append(f"**{n_dif}** diferencia(s) de aforo")
        if n_sl:  partes.append(f"**{n_sl}** especie(s) sin lista en Gallo")
        st.warning(" · ".join(partes))

    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("En lista BYMA",       resumen["total_byma"])
    col_b.metric("OK (coinciden)",       resumen["total_ok"])
    col_c.metric("Diferencias de aforo", n_dif)
    col_d.metric("Sin lista en Gallo",   n_sl)

    if n_lsb:
        st.info(
            f"{n_lsb} especie(s) con lista en Gallo que BYMA ya no acepta "
            "(hoja informacional 'Lista Sin BYMA')."
        )

    # ── Descargas ─────────────────────────────────────────────────────────────
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label="⬇ Descargar DIFERENCIAS AFOROS.xlsx",
            data=result["xlsx_dif"],
            file_name="DIFERENCIAS AFOROS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="caf_dl_dif",
        )
    with col_dl2:
        from datetime import date as _date
        fecha_rc = _date.today().strftime("%d-%m-%y")
        st.download_button(
            label=f"⬇ Descargar Reporte Garantias BYMA {fecha_rc}.xlsx",
            data=result["xlsx_rc"],
            file_name=f"Reporte Garantias BYMA {fecha_rc}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="caf_dl_rc",
        )

    if advertencias:
        with st.expander(f"⚠ {len(advertencias)} advertencia(s) de procesamiento"):
            for adv in advertencias:
                st.markdown(f"- {adv}")
