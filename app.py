import streamlit as st
import pandas as pd
import os
from PIL import Image
import xlwings as xw

# ===============================
# Rutas de archivos
# ===============================
archivo_excel = r"C:\Users\SANTRICH\Desktop\Maykol\Base de datos.xlsm"
carpeta_imagenes = r"C:\Users\SANTRICH\Desktop\Maykol\Python\Imagenes"

# ===============================
# Funciones de carga
# ===============================
@st.cache_data
def cargar_datos():
    df_cualidades = pd.read_excel(archivo_excel, sheet_name="Cualidades")
    df_parametros = pd.read_excel(archivo_excel, sheet_name="Estilos y matrices")
    df_tiempos = pd.read_excel(archivo_excel, sheet_name="Base de datos")
    return df_cualidades, df_parametros, df_tiempos

@st.cache_data
def cargar_imagenes(carpeta):
    extensiones_validas = (".png", ".jpg", ".jpeg", ".PNG", ".JPG", ".JPEG")
    datos = []
    for nombre_archivo in os.listdir(carpeta):
        if nombre_archivo.lower().endswith(extensiones_validas):
            nombre_sin_ext = os.path.splitext(nombre_archivo)[0]
            if nombre_sin_ext.isdigit():
                ruta = os.path.join(carpeta, nombre_archivo)
                orden = int(nombre_sin_ext)
                datos.append({"Orden": orden, "Ruta_Imagen": ruta})
    df_img = pd.DataFrame(datos)
    if not df_img.empty:
        df_img = df_img.sort_values("Orden").reset_index(drop=True)
    return df_img

# ===============================
# Inicializar datos en session_state
# ===============================
if "df_cualidades" not in st.session_state or "df_parametros" not in st.session_state or "df_tiempos" not in st.session_state:
    df_cualidades, df_parametros, df_tiempos = cargar_datos()
    st.session_state["df_cualidades"] = df_cualidades
    st.session_state["df_parametros"] = df_parametros
    st.session_state["df_tiempos"] = df_tiempos
else:
    df_cualidades = st.session_state["df_cualidades"]
    df_parametros = st.session_state["df_parametros"]
    df_tiempos = st.session_state["df_tiempos"]

df_imagenes = cargar_imagenes(carpeta_imagenes)

# ===============================
# Variables de navegación
# ===============================
if "opcion" not in st.session_state:
    st.session_state["opcion"] = "Consultar órdenes"
if "orden_seleccionada" not in st.session_state:
    st.session_state["orden_seleccionada"] = ""

# ===============================
# Sidebar de navegación
# ===============================
st.sidebar.title("Menú principal")
opcion = st.sidebar.radio(
    "Selecciona la acción que deseas realizar:",
    ("Consultar órdenes", "Clasificar prendas", "Consultar tiempos"),
    index=("Consultar órdenes", "Clasificar prendas", "Consultar tiempos").index(st.session_state["opcion"])
)
st.session_state["opcion"] = opcion

# ===============================
# Opción 1: Consultar órdenes
# ===============================
if opcion == "Consultar órdenes":
    st.title("🔍 Consultar órdenes")

    st.subheader("Filtros disponibles")
    num_cols_filtros = 3
    col_filtros = st.columns(num_cols_filtros)

    filtros = {}
    for i, col in enumerate(df_cualidades.columns):
        valores_unicos = df_cualidades[col].dropna().unique()
        with col_filtros[i % num_cols_filtros]:
            if len(valores_unicos) <= 50:
                seleccion = st.multiselect(f"{col}", sorted(valores_unicos))
                if seleccion:
                    filtros[col] = seleccion
            else:
                texto = st.text_input(f"Buscar en {col}")
                if texto:
                    filtros[col] = texto

    # Aplicar filtros
    datos_filtrados = df_cualidades.copy()
    for col, seleccion in filtros.items():
        if isinstance(seleccion, list):
            datos_filtrados = datos_filtrados[datos_filtrados[col].isin(seleccion)]
        else:
            datos_filtrados = datos_filtrados[
                datos_filtrados[col].astype(str).str.contains(seleccion, case=False, na=False)
            ]

    st.subheader("Resultados")
    st.dataframe(datos_filtrados, use_container_width=True)

    if not datos_filtrados.empty:
        st.subheader("Imágenes de las órdenes encontradas")

        num_cols = 4
        cols = st.columns(num_cols)

        for i, (_, fila_datos) in enumerate(datos_filtrados.iterrows()):
            orden = int(fila_datos["Orden"])
            fila_img = df_imagenes[df_imagenes["Orden"] == orden]

            col = cols[i % num_cols]
            with col:
                if not fila_img.empty:
                    ruta_img = fila_img.iloc[0]["Ruta_Imagen"]
                    imagen = Image.open(ruta_img)
                    st.image(imagen, caption=f"Orden {orden}", use_container_width=True)

                    # === Botones centrados debajo de la imagen ===
                    bcol1, bcol2, bcol3 = st.columns([1, 2, 1])  # el del medio más ancho
                    with bcol2:  # centramos los botones
                        c1, c2 = st.columns(2)
                        with c1:
                            if st.button("⏱️", key=f"tiempos_{orden}"):
                                st.session_state["opcion"] = "Consultar tiempos"
                                st.session_state["orden_seleccionada"] = str(orden)
                                st.rerun()
                        with c2:
                            if st.button("✏️", key=f"clasif_{orden}"):
                                st.session_state["opcion"] = "Clasificar prendas"
                                st.session_state["orden_seleccionada"] = str(orden)
                                st.rerun()
                else:
                    st.warning(f"⚠️ No se encontró imagen para la orden {orden}")

# ===============================
# Opción 2: Clasificar prendas
# ===============================
elif opcion == "Clasificar prendas":
    st.title("👕 Clasificar prendas")

    # Orden precargada si viene de otra vista
    numero_orden = st.text_input("Número de orden", value=st.session_state.get("orden_seleccionada", ""))

    # Variable para valores precargados
    valores_form = {}
    valores_existentes = {}
    
    if numero_orden.isdigit():
        orden_int = int(numero_orden)
        fila_existente = df_cualidades[df_cualidades["Orden"].astype(int) == orden_int]
        if not fila_existente.empty:
            valores_existentes = fila_existente.iloc[0].to_dict()
            st.info(f"📂 Orden {orden_int} encontrada, cargando valores existentes...")
        else:
            st.warning(f"🆕 La orden {orden_int} no existe, puedes crearla.")

    # Crear formulario dinámico
    with st.form("form_clasificacion"):
        valores_form["Orden"] = numero_orden

        otros_campos = [c for c in df_cualidades.columns if c != "Orden"]
        for i in range(0, len(otros_campos), 2):
            cols = st.columns(2)
            for j, col in enumerate(otros_campos[i:i+2]):
                with cols[j]:
                    valor_default = valores_existentes.get(col, "")

                    if col in df_parametros.columns:
                        opciones = df_parametros[col].dropna().unique().tolist()
                        opciones = ["N/A"] + opciones

                        # Caso especial: multiselección
                        if col in ["Tipo de tejido", "Apliques"]:
                            valores_previos = []
                            if isinstance(valor_default, str) and valor_default.strip():
                                valores_previos = [v.strip() for v in valor_default.split(",")]

                            seleccion = st.multiselect(
                                col, opciones, default=valores_previos, key=f"form_{col}"
                            )
                            # Guardamos como cadena separada por comas
                            valores_form[col] = ", ".join(seleccion) if seleccion else "N/A"

                        else:
                            if valor_default in opciones:
                                idx = opciones.index(valor_default)
                            else:
                                idx = 0  # N/A por defecto

                            valores_form[col] = st.selectbox(
                                col, opciones, index=idx, key=f"form_{col}"
                            )
                    else:
                        valores_form[col] = st.text_input(
                            col, value=valor_default, key=f"form_{col}"
                        )

        submit = st.form_submit_button("Guardar clasificación")

    # Mostrar imagen de la orden si existe
    if numero_orden.isdigit():
        orden_int = int(numero_orden)
        fila_img = df_imagenes[df_imagenes["Orden"] == orden_int]
        if not fila_img.empty:
            ruta_img = fila_img.iloc[0]["Ruta_Imagen"]
            imagen = Image.open(ruta_img)
            st.image(imagen, caption=f"Orden {orden_int}", use_container_width=True)

    # Guardar cambios
    if submit:
        if not numero_orden.isdigit():
            st.error("❌ Debes ingresar un número de orden válido")
        else:
            orden_int = int(numero_orden)

            try:
                # Abrir Excel con xlwings
                app = xw.App(visible=False)
                wb = app.books.open(archivo_excel)
                ws = wb.sheets["Cualidades"]

                headers = ws.range("A1").expand("right").value
                ultima_fila = ws.range("A1").expand("down").last_cell.row

                df_excel = pd.DataFrame(
                    ws.range("A1").expand().value[1:], 
                    columns=headers
                )

                if orden_int in df_excel["Orden"].astype(int).values:
                    fila_excel = df_excel.index[df_excel["Orden"].astype(int) == orden_int][0]
                    for col, val in valores_form.items():
                        if col in headers:
                            ws.range((fila_excel + 2, headers.index(col) + 1)).value = val
                    st.success(f"🔄 Orden {orden_int} actualizada con éxito")
                else:
                    nueva_fila = [valores_form.get(col, "") for col in headers]
                    ws.range((ultima_fila + 1, 1)).value = nueva_fila
                    st.success(f"🆕 Orden {orden_int} creada con éxito")

                wb.save()
                wb.close()
                app.quit()

                # 🔄 Recargar datos
                cargar_datos.clear()
                df_cualidades, df_parametros = cargar_datos()
                st.session_state["df_cualidades"] = df_cualidades
                st.session_state["df_parametros"] = df_parametros

                st.info("💾 Cambios guardados y base de datos recargada")

            except Exception as e:
                st.error(f"❌ Error al guardar en Excel con xlwings: {e}")


# ===============================
# Opción 3: Consultar tiempos
# ===============================
elif opcion == "Consultar tiempos":
    st.title("⏱️ Consultar tiempos")

    # Orden precargada si viene de otra vista
    orden = st.text_input("Número de orden para consulta de tiempos", value=st.session_state.get("orden_seleccionada", ""))

    if orden:
        try:
            # Filtrar por la orden
            datos_orden = df_tiempos[df_tiempos["Orden"].astype(str) == str(orden)]

            if not datos_orden.empty:
                # Mostrar información general
                descripcion = datos_orden["DESCRIPCIÓN PRODUCTO"].iloc[0]
                cliente = datos_orden["CLIENTE"].iloc[0]

                st.subheader(f"📌 Orden {orden}")
                st.write(f"**📦 Producto:** {descripcion}")
                st.write(f"**📌 Cliente:** {cliente}")

                # Agrupar por sección y sumar tiempos
                tiempos_resumen = (
                    datos_orden.groupby("SECCIÓN")["TIEMPO"]
                    .sum()
                    .reset_index()
                )

                # Mostrar tabla resumen
                st.subheader("📊 Resumen de tiempos por sección")
                st.dataframe(tiempos_resumen)

                # Mostrar tabla detallada por operación
                st.write("🔹 **Detalle de operaciones**")
                st.dataframe(
                    datos_orden[["SECCIÓN", "DESCRIPCIÓN OPERACIÓN", "TIEMPO"]]
                )

                # Tiempo total
                tiempo_total = tiempos_resumen["TIEMPO"].sum()
                st.success(f"⏱️ Tiempo total de la orden {orden}: {tiempo_total:.2f}")

                # Gráfico de barras con plotly
                import plotly.express as px
                fig = px.bar(
                    tiempos_resumen,
                    x="SECCIÓN",
                    y="TIEMPO",
                    text="TIEMPO",
                    title=f"Tiempos por sección - Orden {orden}",
                    labels={"SECCIÓN": "Sección", "TIEMPO": "Tiempo (min)"},
                    color="SECCIÓN"
                )
                fig.update_traces(texttemplate="%{text:.2f}", textposition="outside")
                fig.update_layout(
                    xaxis_title="Sección",
                    yaxis_title="Tiempo (minutos)",
                    uniformtext_minsize=8,
                    uniformtext_mode="hide",
                    plot_bgcolor="white"
                )
                st.plotly_chart(fig, use_container_width=True)

                # ===============================
                # Mostrar imágenes de la orden
                # ===============================
                st.subheader("🖼️ Imágenes de la orden")
                datos_filtrados = df_cualidades[df_cualidades["Orden"].astype(str) == str(orden)]

                if not datos_filtrados.empty:
                    num_cols = 3  # número de imágenes por fila
                    cols = st.columns(num_cols)

                    for i, (_, fila_datos) in enumerate(datos_filtrados.iterrows()):
                        orden_val = int(fila_datos["Orden"])
                        fila_img = df_imagenes[df_imagenes["Orden"] == orden_val]

                        if not fila_img.empty:
                            ruta_img = fila_img.iloc[0]["Ruta_Imagen"]
                            imagen = Image.open(ruta_img)

                            col = cols[i % num_cols]  # asignar a columna
                            with col:
                                st.image(imagen, caption=f"Orden {orden_val}", use_container_width=True)
                        else:
                            st.warning(f"⚠️ No se encontró imagen para la orden {orden_val}")
                else:
                    st.warning(f"⚠️ No se encontraron registros en 'Cualidades' para la orden {orden}")

            else:
                st.warning(f"No se encontraron tiempos para la orden {orden}")

        except Exception as e:
            st.error(f"Error al consultar los tiempos: {e}")
