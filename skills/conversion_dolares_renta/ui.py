import streamlit as st
import traceback


def render():
    st.title("Conversión Dólares Renta")
    st.markdown(
        "Lee el **Diario.xlsx** con los saldos de dólares renta por comitente y genera "
        "los archivos de conversión para importar en Gallo y en NASDAQ BYMA."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Archivo del día — subir acá:**
        - **Diario.xlsx** — planilla con saldos por comitente.
          - Fila 1: `Plantilla diaria` + fecha del día en celda **D1**
          - Fila 3: encabezados (`Ctte | DÓLAR USA a MEP | DÓLAR USA a USD CABLE | DÓLAR DLR`)
          - Fila 4+: datos (un comitente por fila; puede tener valor en una, dos o tres columnas)

        **Archivos generados:**
        - `GALLO Conversion Especie a Moneda 7K DD-MM-AA.xls` — comitentes con DÓLAR USA a MEP (columna B)
        - `GALLO Conversion Especie a Moneda 7K USD CABLE DD-MM-AA.xls` — comitentes con DÓLAR USA a USD CABLE (columna C)
        - `GALLO Conversion Especie a Moneda 10K DD-MM-AA.xls` — comitentes con DÓLAR DLR (columna D)
        - `Archivo masivo - transferencias de efectivo DD-MM-AA.ICT` — transferencias para NASDAQ BYMA

        **Estructura del ICT:**
        - Col B (7K MEP): `USD-EXTERNO` → cuenta `233/1`
        - Col C (7K Cable): `USD-EXTERNO` → cuenta `70233/10000`
        - Col D (10K): `USD-LOCAL` → cuenta `233/1`
        - Referencias: `ConvAAMMDD####` — el contador persiste durante la sesión

        **Contadores Conv:** persisten durante la sesión para evitar referencias duplicadas
        si se ejecuta más de una vez el mismo día. Usar el botón de reset solo al comenzar el día.
        """)

    # ── Upload ────────────────────────────────────────────────────────────────
    st.subheader("Archivo del día")
    diario_file = st.file_uploader(
        "Diario.xlsx",
        type=["xlsx", "xls"],
        key="cdr_diario",
        help="Planilla diaria con saldos de dólares renta por comitente"
    )

    # ── Contador Conv ─────────────────────────────────────────────────────────
    st.divider()
    if "cdr_counter" not in st.session_state:
        st.session_state["cdr_counter"] = {}

    counter = st.session_state["cdr_counter"]
    col_ctr, col_btn = st.columns([3, 1])
    with col_ctr:
        if counter:
            st.caption(
                "Contadores activos: "
                + ", ".join(f"{k}→{v:04d}" for k, v in sorted(counter.items()))
            )
        else:
            st.caption("Contador en 0 — el próximo proceso arranca desde Conv…0001.")
    with col_btn:
        if st.button("Reiniciar contador", key="cdr_reset"):
            st.session_state["cdr_counter"] = {}
            st.rerun()

    st.divider()

    if not diario_file:
        st.info("Subí el Diario.xlsx para habilitar la generación.")
        return

    # ── Ejecución ─────────────────────────────────────────────────────────────
    if st.button("Generar archivos", type="primary", use_container_width=True):
        with st.spinner("Procesando Diario..."):
            try:
                from skills.conversion_dolares_renta.logic import generar_conversion
                outputs, new_counter, resumen, advertencias = generar_conversion(
                    diario_file=diario_file,
                    counter_state=st.session_state["cdr_counter"],
                )
                st.session_state["cdr_counter"] = new_counter

                # ── Estado global ─────────────────────────────────────────────
                if resumen["verificacion_ok"]:
                    st.success(
                        f"✓ Archivos generados — {resumen['fecha']}  |  "
                        f"{resumen['n_ict']} línea(s) ICT  |  "
                        f"Refs: {resumen['ref_primera']} – {resumen['ref_ultima']}"
                    )
                else:
                    st.warning(
                        f"⚠ Archivos generados con advertencias — {resumen['fecha']}"
                    )

                if resumen["contador_continuado"]:
                    st.info(
                        f"Contador continuado desde ejecución anterior "
                        f"(inicio en {resumen['inicio_contador']:04d})."
                    )

                # ── Métricas ──────────────────────────────────────────────────
                col_a, col_b, col_c, col_d = st.columns(4)
                col_a.metric("7K MEP",       resumen["n_7k"],    help="Comitentes col B")
                col_b.metric("7K USD Cable", resumen["n_cable"], help="Comitentes col C")
                col_c.metric("10K",          resumen["n_10k"],   help="Comitentes col D")
                col_d.metric("Líneas ICT",   resumen["n_ict"])

                # ── Descargas ─────────────────────────────────────────────────
                st.subheader("Descargas")

                # Fila 1: los tres XLS de Gallo
                col1, col2, col3 = st.columns(3)
                with col1:
                    if outputs["gallo_7k"]:
                        content, fname = outputs["gallo_7k"]
                        st.download_button(
                            label=f"⬇ {fname}",
                            data=content,
                            file_name=fname,
                            mime="application/vnd.ms-excel",
                            use_container_width=True,
                            key="cdr_dl_7k",
                        )
                    else:
                        st.caption("Sin datos 7K MEP")

                with col2:
                    if outputs["gallo_7k_cable"]:
                        content, fname = outputs["gallo_7k_cable"]
                        st.download_button(
                            label=f"⬇ {fname}",
                            data=content,
                            file_name=fname,
                            mime="application/vnd.ms-excel",
                            use_container_width=True,
                            key="cdr_dl_7k_cable",
                        )
                    else:
                        st.caption("Sin datos 7K Cable")

                with col3:
                    if outputs["gallo_10k"]:
                        content, fname = outputs["gallo_10k"]
                        st.download_button(
                            label=f"⬇ {fname}",
                            data=content,
                            file_name=fname,
                            mime="application/vnd.ms-excel",
                            use_container_width=True,
                            key="cdr_dl_10k",
                        )
                    else:
                        st.caption("Sin datos 10K")

                # Fila 2: ICT — ancho completo
                if outputs["ict"]:
                    content, fname = outputs["ict"]
                    st.download_button(
                        label=f"⬇ {fname}  ({resumen['n_ict']} transferencias)",
                        data=content,
                        file_name=fname,
                        mime="text/plain",
                        use_container_width=True,
                        key="cdr_dl_ict",
                    )

                # ── Advertencias ──────────────────────────────────────────────
                if advertencias:
                    with st.expander(f"⚠ {len(advertencias)} advertencia(s)"):
                        for adv in advertencias:
                            st.markdown(f"- {adv}")

            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())
