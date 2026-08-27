import os
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Evaluación Fisioterapia - Adulto Mayor",
    page_icon="👴",
    layout="wide",
)

EXCEL_FILE = "base para estudio adulto mayor.xlsx"


def load_or_create_excel():
  if not os.path.exists(EXCEL_FILE):
    # Crear un archivo excel con las 4 pestañas y estructura base si no existe
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
      pd.DataFrame(
          columns=[
              "NOMBRE",
              "DOCUMENTO",
              "EDAD",
              "GENERO",
              "ESTRATO",
              "REGIMEN DE SALUD",
              "EPS",
              "FUENTE DE INGRESOS",
              "ESTADO CIVIL",
              "NIVEL EDUCATIVO",
              "TIEMPO DE INSTITUCIONALIZACION",
              "MOTIVO DE INSTITUCIONALIZACION",
              "LUGAR DE PROCEDENCIA",
              "RED DE APOYO",
              "HOGAR GERIATRICO",
          ]
      ).to_excel(writer, sheet_name="DATOS SOCIODEMOGRAFICOS", index=False)

      pd.DataFrame(
          columns=[
              "DOCUMENTO",
              "NUMERO DE COMORBILIDADES",
              "CUALES",
              "SINDROME GERIATRICO",
              "CAIDAS",
              "FARMACOS",
          ]
      ).to_excel(writer, sheet_name="VARIABLES CLINICAS", index=False)

      pd.DataFrame(
          columns=[
              "DOCUMENTO",
              "FUERZA MMII",
              "FUERZA MMSS",
              "FLEXIBILIDAD MMII",
              "FLEXIBILIDAD MMSS",
              "EQUILIBRIO DINAMICO",
              "CAPACIDAD AEROBICA",
          ]
      ).to_excel(writer, sheet_name="SENIOR FITNESS TEST", index=False)

      pd.DataFrame(
          columns=[
              "DOCUMENTO",
              "FUNCION EJECUTIVA",
              "MEMORIA INMEDIATA",
              "FLUIDEZ",
              "ORIENTACION",
              "CALCULO",
              "ABSTRACCION",
              "MEMORIA DIFERENCIAL",
              "PERCEPCION VISUAL",
              "DENOMINACION",
              "ATENCION",
              "NUMERACION",
              "FUNCION COGNTIVA GLOBAL",
              "CLASIFICACION COGNITIVA",
          ]
      ).to_excel(writer, sheet_name="MOCA-B", index=False)


load_or_create_excel()

st.title("🧓 Sistema de Evaluación Fisioterapéutica en Adultos Mayores")
st.markdown(
    "Herramienta digital para la recolección de datos clínicos, sociodemográficos"
    " y funcionales sin papel."
)

# Menú lateral de navegación
menu = st.sidebar.selectbox(
    "Menú de Navegación",
    [
        "Registrar Evaluación",
        "Ver Registros Guardados",
    ],
)

if menu == "Registrar Evaluación":
  st.header("Formulario de Evaluación Integral")

  with st.form("evaluation_form"):
    # 1. DATOS SOCIODEMOGRAFICOS
    st.subheader("1. Datos Sociodemográficos")
    col1, col2, col3 = st.columns(3)
    with col1:
      nombre = st.text_input("Nombre Completo")
      documento = st.text_input("Documento de Identidad")
      edad = st.number_input("Edad", min_value=50, max_value=120, value=70)
      genero = st.selectbox(
          "Género", ["Masculino", "Femenino", "Otro", "No especificado"]
      )
      estrato = st.selectbox("Estrato Socioeconómico", ["1", "2", "3", "4", "5", "6", "N/A"])
    with col2:
      regimen = st.selectbox("Régimen de Salud", ["Contributivo", "Subsidiado", "Especial", "Ninguno"])
      eps = st.text_input("EPS")
      fuente_ingresos = st.selectbox(
          "Fuente de Ingresos",
          ["Pensión", "Salario", "Familia", "Subsidio/Beneficio", "Ninguna"],
      )
      estado_civil = st.selectbox(
          "Estado Civil",
          ["Soltero/a", "Casado/a", "Unión Libre", "Viudo/a", "Divorciado/a"],
      )
      nivel_educativo = st.selectbox(
          "Nivel Educativo",
          [
              "Ninguno",
              "Primaria",
              "Secundaria",
              "Técnico/Tecnológico",
              "Universitario",
          ],
      )
    with col3:
      tiempo_inst = st.text_input("Tiempo de Institucionalización")
      motivo_inst = st.text_input("Motivo de Institucionalización")
      lugar_proc = st.text_input("Lugar de Procedencia")
      red_apoyo = st.selectbox(
          "Red de Apoyo", ["Familiar", "Institucional", "Amigos", "Ninguna"]
      )
      hogar_geriatrico = st.text_input("Hogar Geriátrico")

    st.markdown("---")

    # 2. VARIABLES CLINICAS
    st.subheader("2. Variables Clínicas")
    col4, col5 = st.columns(2)
    with col4:
      num_comorbilidades = st.number_input(
          "Número de Comorbilidades", min_value=0, max_value=15, value=0
      )
      cuales_com = st.text_area("¿Cuáles comorbilidades?")
      sindrome_geriatrico = st.text_input("Síndrome Geriátrico")
    with col5:
      caidas = st.text_input("Antecedente de Caídas")
      farmacos = st.text_area("Fármacos (Medicamentos)")

    st.markdown("---")

    # 3. SENIOR FITNESS TEST
    st.subheader("3. Senior Fitness Test (Funcionalidad)")
    col6, col7 = st.columns(2)
    with col6:
      fuerza_mmii = st.text_input(
          "Fuerza MMII (Ej. Sentarse y levantarse de la silla)"
      )
      fuerza_mmss = st.text_input("Fuerza MMSS (Ej. Flexión de codo)")
      flexibilidad_mmii = st.text_input(
          "Flexibilidad MMII (Ej. Sentado y alcanzar los dedos del pie)"
      )
    with col7:
      flexibilidad_mmss = st.text_input(
          "Flexibilidad MMSS (Ej. Juntar las manos tras la espalda)"
      )
      equilibrio_dinamico = st.text_input(
          "Equilibrio Dinámico (Ej. Timed Up and Go)"
      )
      capacidad_aerobica = st.text_input(
          "Capacidad Aeróbica (Ej. Marcha de 6 minutos)"
      )

    st.markdown("---")

    # 4. MOCA-B (Cognitivo)
    st.subheader("4. Test MOCA-B (Evaluación Cognitiva)")
    col8, col9, col10 = st.columns(3)
    with col8:
      func_ejecutiva = st.text_input("Función Ejecutiva")
      memoria_inmediata = st.text_input("Memoria Inmediata")
      fluidez = st.text_input("Fluidez Verbal")
      orientacion = st.text_input("Orientación")
    with col9:
      calculo = st.text_input("Cálculo")
      abstraccion = st.text_input("Abstracción")
      memoria_diferencial = st.text_input("Memoria Diferencial")
      percepcion_visual = st.text_input("Percepción Visual")
    with col10:
      denominacion = st.text_input("Denominación")
      atencion = st.text_input("Atención")
      numeracion = st.text_input("Numeración")
      func_global = st.text_input("Función Cognitiva Global")
      clasificacion_cognitiva = st.text_input("Clasificación Cognitiva")

    submit_button = st.form_submit_button(
        label="Guardar Evaluación en Excel"
    )

    if submit_button:
      if not documento:
        st.warning(
            "Por favor, ingrese al menos el documento de identidad para"
            " registrar los datos."
        )
      else:
        # Preparar dataframes para anexar
        df_socio = pd.DataFrame([{
            "NOMBRE": nombre,
            "DOCUMENTO": documento,
            "EDAD": edad,
            "GENERO": genero,
            "ESTRATO": estrato,
            "REGIMEN DE SALUD": regimen,
            "EPS": eps,
            "FUENTE DE INGRESOS": fuente_ingresos,
            "ESTADO CIVIL": estado_civil,
            "NIVEL EDUCATIVO": nivel_educativo,
            "TIEMPO DE INSTITUCIONALIZACION": tiempo_inst,
            "MOTIVO DE INSTITUCIONALIZACION": motivo_inst,
            "LUGAR DE PROCEDENCIA": lugar_proc,
            "RED DE APOYO": red_apoyo,
            "HOGAR GERIATRICO": hogar_geriatrico,
        }])

        df_clinica = pd.DataFrame([{
            "DOCUMENTO": documento,
            "NUMERO DE COMORBILIDADES": num_comorbilidades,
            "CUALES": cuales_com,
            "SINDROME GERIATRICO": sindrome_geriatrico,
            "CAIDAS": caidas,
            "FARMACOS": farmacos,
        }])

        df_fitness = pd.DataFrame([{
            "DOCUMENTO": documento,
            "FUERZA MMII": fuerza_mmii,
            "FUERza MMSS": fuerza_mmss,
            "FLEXIBILIDAD MMII": flexibilidad_mmii,
            "FLEXIBILIDAD MMSS": flexibilidad_mmss,
            "EQUILIBRIO DINAMICO": equilibrio_dinamico,
            "CAPACIDAD AEROBICA": capacidad_aerobica,
        }])

        df_moca = pd.DataFrame([{
            "DOCUMENTO": documento,
            "FUNCION EJECUTIVA": func_ejecutiva,
            "MEMORIA INMEDIATA": memoria_inmediata,
            "FLUIDEZ": fluidez,
            "ORIENTACION": orientacion,
            "CALCULO": calculo,
            "ABSTRACCION": abstraccion,
            "MEMORIA DIFERENCIAL": memoria_diferencial,
            "PERCEPCION VISUAL": percepcion_visual,
            "DENOMINACION": denominacion,
            "ATENCION": atencion,
            "NUMERACION": numeracion,
            "FUNCION COGNTIVA GLOBAL": func_global,
            "CLASIFICACION COGNITIVA": clasificacion_cognitiva,
        }])

        # Guardar conservando las hojas existentes
        with pd.ExcelWriter(
            EXCEL_FILE, mode="a", engine="openpyxl", if_sheet_exists="overlay"
        ) as writer:
          for sheet_name, new_df in zip(
              [
                  "DATOS SOCIODEMOGRAFICOS",
                  "VARIABLES CLINICAS",
                  "SENIOR FITNESS TEST",
                  "MOCA-B",
              ],
              [df_socio, df_clinica, df_fitness, df_moca],
          ):
            try:
              existing_df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
              combined_df = pd.concat(
                  [existing_df, new_df], ignore_index=True
              )
            except Exception:
              combined_df = new_df
            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)

        st.success(
            "¡Evaluación guardada exitosamente en el archivo Excel del"
            " proyecto!"
        )

elif menu == "Ver Registros Guardados":
  st.header("Base de Datos Actualizada")
  if os.path.exists(EXCEL_FILE):
    sheet_to_show = st.selectbox(
        "Seleccione la sección a visualizar",
        [
            "DATOS SOCIODEMOGRAFICOS",
            "VARIABLES CLINICAS",
            "SENIOR FITNESS TEST",
            "MOCA-B",
        ],
    )
    df_view = pd.read_excel(EXCEL_FILE, sheet_name=sheet_to_show)
    st.dataframe(df_view)

    # Botón para descargar el archivo actualizado
    with open(EXCEL_FILE, "rb") as f:
      st.download_button(
          label="Descargar Excel Completo",
          data=f,
          file_name="base_adulto_mayor_actualizada.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
  else:
    st.info("Aún no hay registros guardados.")