import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Control Márgenes Gara BYMA")
    st.markdown(
        "Compara la circular de aforos BYMA (PDF) con el maestro de especies de Gallo (ESPECIES.XLS). "
        "Detecta diferencias de lista/aforo y especies aceptadas por BYMA que no están dadas de alta en Gallo."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Archivos de referencia** (se toman automáticamente de Archivos Compartidos):
        - **ESPECIES.XLS** — maestro de especies: código CVSA y Lista asignada (col F)
        - **PDF de aforos BYMA** — circular con las especies aceptadas como garantía y su aforo (%)

        **Output — Excel DIFERENCIAS AFOROS.xlsx:**
        - Hoja **Diferencias Aforo**: especies cuya Lista en Gallo no coincide con el aforo del PDF.
          Columna *Aforo Sailing* en rojo (valor actual), *Lista Sugerida* en verde (corrección sugerida).
        - Hoja **No Encontradas** *(si aplica)*: especies del PDF que no están en el maestro de Gallo.

        **Output — FALTANTE en GALLO.txt** *(si aplica)*: listado de especies para dar de alta en Gallo.

        **Tablas de referencia Lista → Aforo:**

        | Segmento | Listas | Aforos |
        |---|---|---|
        | Renta Variable | 1–8 | 85%–30% |
        | Renta Fija Públicos | 11–17 | 90%–60% |
        | Renta Fija Privados | 22–27 | 85%–60% |
        | Letras y Bonos del Tesoro | 85 / 90 / 95 | 85% / 90% / 95% |
        """)

    # ── Inputs (ambos desde Archivos Compartidos) ─────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        especies_file = shared_or_upload(
            "shared_especies", "Maestro de especies (ESPECIES.XLS)",
            ["xls", "xlsx"], "caf_especies"
        )
    with col2:
        pdf_file = shared_or_upload(
            "shared_pdf_aforos", "PDF de aforos BYMA",
            ["pdf"], "caf_pdf"
        )

    st.divider()

    # ── Validación ────────────────────────────────────────────────────────────
    faltantes_inp = []
    if not especies_file: faltantes_inp.append("ESPECIES.XLS")
    if not pdf_file:      faltantes_inp.append("PDF de aforos BYMA")

    if faltantes_inp:
        st.info(f"Cargá desde Archivos Compartidos: {', '.join(faltantes_inp)}")
        return

    # ── Ejecución ─────────────────────────────────────────────────────────────
    if st.button("Ejecutar control", type="primary", use_container_width=True):
        with st.spinner("Procesando PDF y maestro de especies..."):
            try:
                from skills.control_aforos_byma.logic import generar_control
                xlsx_bytes, txt_faltantes, resumen, advertencias = generar_control(
                    especies_file, pdf_file
                )

                # Mensaje de estado
                n_dif  = resumen["diferencias"]
                n_falt = resumen["faltantes"]
                if n_dif == 0 and n_falt == 0:
                    st.success("✓ Sin diferencias ni faltantes — maestro alineado con circular BYMA.")
                else:
                    partes = []
                    if n_dif:  partes.append(f"**{n_dif}** diferencia(s) de aforo")
                    if n_falt: partes.append(f"**{n_falt}** especie(s) faltante(s) en Gallo")
                    st.warning(" · ".join(partes))

                # Métricas
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("Especies en PDF",       resumen["total_pdf"])
                col_b.metric("Diferencias de aforo",  n_dif)
                col_c.metric("Faltantes en Gallo",    n_falt)

                # Descargas
                st.download_button(
                    label="⬇ Descargar DIFERENCIAS AFOROS.xlsx",
                    data=xlsx_bytes,
                    file_name="DIFERENCIAS AFOROS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                if txt_faltantes:
                    st.download_button(
                        label="⬇ Descargar FALTANTE en GALLO.txt",
                        data=txt_faltantes.encode("utf-8"),
                        file_name="FALTANTE en GALLO.txt",
                        mime="text/plain",
                        use_container_width=True,
                    )

                # Advertencias de procesamiento
                if advertencias:
                    with st.expander(f"⚠ {len(advertencias)} advertencia(s) de procesamiento"):
                        for adv in advertencias:
                            st.markdown(f"- {adv}")

            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())
