import streamlit as st
import pandas as pd
import numpy as np
import io
import unicodedata
import calendar

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA WEB
# ==========================================
st.set_page_config(page_title="Plataforma Suplencias - CAPJ", page_icon="⚖️", layout="centered")

# ==========================================
# UTILIDADES DE FORMATO
# ==========================================
def normalizar_texto(texto):
    if not isinstance(texto, str): return texto
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto.upper().replace('_', ' ')

# ==========================================
# INTERFAZ DE USUARIO (FRONTEND)
# ==========================================
st.title("⚖️ Plataforma Automática de Suplencias - CAPJ")
st.markdown("Cargue los 5 archivos Excel requeridos y presione **Procesar y Formatear** para generar la matriz de remuneraciones final.")

# Organización de la interfaz en dos columnas
col1, col2 = st.columns(2)
with col1:
    up_suplencia = st.file_uploader("1. CertificadoSuplencia", type=['xls', 'xlsx'])
    up_zonas = st.file_uploader("3. Asignacion Zona", type=['xls', 'xlsx'])
    up_sueldos = st.file_uploader("5. Sueldos Base", type=['xls', 'xlsx'])
with col2:
    up_maestro = st.file_uploader("2. Maestro Histórico", type=['xls', 'xlsx'])
    up_banrem = st.file_uploader("4. BanRemExport", type=['xls', 'xlsx'])

# ==========================================
# MOTOR DE PROCESAMIENTO (BACKEND)
# ==========================================
if st.button("⚙️ Procesar y Formatear", type="primary", use_container_width=True):
    if not (up_suplencia and up_maestro and up_zonas and up_banrem and up_sueldos):
        st.error("❌ Error: Faltan archivos por cargar. Por favor suba las 5 bases de datos.")
    else:
        with st.spinner("⏳ Procesando datos y calculando devengos..."):
            try:
                # Carga de datos directa desde la interfaz web
                df_banrem = pd.read_excel(up_banrem, skiprows=6, usecols="A:Q")
                df_cert_suplencia = pd.read_excel(up_suplencia)
                df_maestro_raw = pd.read_excel(up_maestro, usecols="A:E, K, L, N, AB:AE, AI:AN, BA")
                df_zona_tabla = pd.read_excel(up_zonas)
                df_sueldos_ref = pd.read_excel(up_sueldos)

                # Fase 1: Comuna Unidad y Llaves
                if 'Unidad' in df_banrem.columns:
                    comunas_extraidas = df_banrem['Unidad'].astype(str).str.split().str[-1]
                    df_banrem.insert(df_banrem.columns.get_loc('Unidad') + 1, 'Comuna_Unidad', comunas_extraidas)

                reemplazos_comunas = {'ANGELES': 'LOS ANGELES', 'BARBARA': 'SANTA BARBARA', 'JUANA': 'SANTA JUANA'}
                for col_c in ['Comuna_Unidad', 'COMUNA UNIDAD']:
                    if col_c in df_banrem.columns:
                        df_banrem[col_c] = df_banrem[col_c].astype(str).str.upper().replace(reemplazos_comunas)
                    if col_c in df_maestro_raw.columns:
                        df_maestro_raw[col_c] = df_maestro_raw[col_c].astype(str).str.upper().replace(reemplazos_comunas)

                def limpiar_run_fecha(df, c_run, c_fec):
                    df[c_run] = df[c_run].astype(str).str.split('.').str[0].str.replace('-', '').str.strip()
                    df[c_fec] = pd.to_datetime(df[c_fec]).dt.strftime('%Y-%m-%d')
                    df['LLAVE_CONCATENADA'] = df[c_run] + "_" + df[c_fec]
                    return df

                df_banrem = limpiar_run_fecha(df_banrem, 'RUN', 'Fecha Inicio')
                df_cert_suplencia = limpiar_run_fecha(df_cert_suplencia, 'RUN', 'Fecha Inicio')
                
                resultado_fase2 = df_banrem.merge(df_cert_suplencia[['LLAVE_CONCATENADA', 'Grado']], on='LLAVE_CONCATENADA', how='left')

                # Fase 2: Maestro
                df_maestro_raw['RUN'] = df_maestro_raw['RUN'].astype(str).str.split('.').str[0].str.replace('-', '').str.strip()
                df_maestro = df_maestro_raw[df_maestro_raw['ESTADO'].isin(['Vigente', 'Prorrogado'])]
                resultado_fase2 = resultado_fase2.merge(df_maestro, on='RUN', how='inner')

                # Fase 3: Zonas %
                mapa_zonas = dict(zip(df_zona_tabla['COMUNA'].astype(str).str.strip().str.upper(), df_zona_tabla['ASIGNACION DE ZONA']))
                cols_com = [c for c in resultado_fase2.columns if 'COMUNA' in c.upper()]
                for c in cols_com:
                    vals = resultado_fase2[c].astype(str).str.strip().str.upper().map(mapa_zonas)
                    resultado_fase2.insert(resultado_fase2.columns.get_loc(c) + 1, f'Zona_Asig_{c}', vals)

                # Fase 4: Sueldos
                mapa_sueldos = dict(zip(df_sueldos_ref['GRADOS'].astype(str).str.strip(), df_sueldos_ref['SUELDO BASE']))
                for g_col in [c for c in resultado_fase2.columns if 'GRADO' in c.upper()]:
                    vals = resultado_fase2[g_col].astype(str).str.split('.').str[0].str.strip().map(mapa_sueldos)
                    resultado_fase2.insert(resultado_fase2.columns.get_loc(g_col) + 1, f'Sueldo_Base_{g_col}', vals)

                # Fase 5: Montos Mensuales
                z_tele_temp = resultado_fase2['Zona_Asig_Comuna'].fillna(resultado_fase2['Zona_Asig_Comuna_Unidad'])
                p_min_sup = np.minimum(resultado_fase2['Zona_Asig_Comuna_Unidad'], z_tele_temp)
                m_sup = (resultado_fase2['Sueldo_Base_Grado'] * (p_min_sup / 100)).round(0)

                c_s_ori = [c for c in resultado_fase2.columns if 'Sueldo_Base' in c and c != 'Sueldo_Base_Grado'][0]
                c_z_ori = [c for c in resultado_fase2.columns if 'Zona_Asig' in c and c not in ['Zona_Asig_Comuna_Unidad', 'Zona_Asig_Comuna']][0]
                m_ori = (resultado_fase2[c_s_ori] * (resultado_fase2[c_z_ori] / 100)).round(0)

                resultado_fase2.insert(resultado_fase2.columns.get_loc('Zona_Asig_Comuna_Unidad') + 1, 'Monto_Asignacion_Zona_Suplencia', m_sup)
                resultado_fase2.insert(resultado_fase2.columns.get_loc(c_z_ori) + 1, 'Monto_Asignacion_Zona_Origen', m_ori)

                # Fase 6 y 7: Cálculos de fechas y días
                resultado_fase2 = resultado_fase2.rename(columns={'Fecha Inicio.1': 'Fecha Inicio Suplencia', 'Fecha Término.1': 'Fecha Término Suplencia'})
                
                cols_fec_calc = ['Fecha Inicio Suplencia', 'Fecha Término Suplencia', 'Fecha Inicio', 'Fecha Término', 'Fecha de Certificación']
                for cf in cols_fec_calc:
                    if cf in resultado_fase2.columns:
                        resultado_fase2[cf] = pd.to_datetime(resultado_fase2[cf], errors='coerce')

                long_suplencia = (resultado_fase2['Fecha Término Suplencia'] - resultado_fase2['Fecha Inicio Suplencia']).dt.days + 1
                resultado_fase2['día o mes'] = np.where(long_suplencia >= 90, 'MES', 'DIA')

                def calcular_dias_automatico(fila):
                    if pd.isna(fila['Fecha Inicio']) or pd.isna(fila['Fecha Término']): return 0
                    db = (fila['Fecha Término'] - fila['Fecha Inicio']).days
                    return (db + 1) if fila['día o mes'] == 'DIA' else db

                resultado_fase2['días'] = resultado_fase2.apply(calcular_dias_automatico, axis=1)

                resultado_fase2['Zona_Proporcional_Pagar_Suplencia'] = (resultado_fase2['Monto_Asignacion_Zona_Suplencia'] / 30 * resultado_fase2['días']).round(0)
                resultado_fase2['Zona_Proporcional_Pagar_Origen'] = (resultado_fase2['Monto_Asignacion_Zona_Origen'] / 30 * resultado_fase2['días']).round(0)
                resultado_fase2['Diferencia_Asig_Zona'] = resultado_fase2['Zona_Proporcional_Pagar_Suplencia'] - resultado_fase2['Zona_Proporcional_Pagar_Origen']

                # Formato Fechas
                for cf in cols_fec_calc:
                    if cf in resultado_fase2.columns:
                        resultado_fase2[cf] = resultado_fase2[cf].dt.strftime('%d/%m/%Y')

                # Renombre
                dict_renombre = {
                    'Zona_Asig_Comuna_Unidad': 'ASIG. ZONA SUPLENCIA',
                    'Monto_Asignacion_Zona_Suplencia': 'MONTO ASIG. ZONA SUPLENCIA',
                    'Zona_Proporcional_Pagar_Suplencia': 'PROPORCIONAL ZONA SUPLENCIA',
                    'Fecha de Certificación': 'FECHA CERTIFICACION',
                    'Comuna_Unidad': 'COMUNA SUPLENCIA',
                    'Comuna': 'COMUNA TT SUPLENCIA',
                    'COMUNA UNIDAD' : 'COMUNA ORIGEN',
                    'Fecha Inicio': 'FECHA INICIO CERT. SUP.',
                    'Fecha Término': 'FECHA TERMINO CERT. SUP.',
                    'Fecha Inicio Suplencia': 'FECHA INICIO SUPLENCIA',
                    'Fecha Término Suplencia': 'FECHA TERMINO SUPLENCIA',
                    'Grado': 'GRADO SUPLENCIA',
                    'GRADO' : 'GRADO ORIGEN',
                    'CARGO' : 'CARGO ORIGEN',
                    'Sueldo_Base_GRADO' : 'SUELDO BASE ORIGEN',
                    'Sueldo_Base_Grado': 'SUELDO BASE SUPLENCIA',
                    'Zona_Asig_COMUNA UNIDAD': 'ASIG. ZONA COMUNA ORIGEN',
                    'Monto_Asignacion_Zona_Origen': 'MONTO ASIG. ZONA ORIGEN',
                    'Zona_Proporcional_Pagar_Origen': 'PROPORCIONAL ZONA ORIGEN',
                    'UNIDAD LABORAL': 'UNIDAD LABORAL ORIGEN',
                    'Unidad': 'UNIDAD LABORAL SUPLENCIA',
                    'Diferencia_Asig_Zona': 'DIFERENCIA ASIG. ZONA'
                }
                resultado_fase2 = resultado_fase2.rename(columns=dict_renombre)

                # Limpieza y Orden
                cols_quitar = ["ESCALA DE SUELDO", "ASIENTO UNIDAD", "TIPO UNIDAD LABORAL", "Tipo", "Estado",
                               "FECHA INGRESO AL SERVICIO", "FECHA INGRESO A LA PLANTA", "ORGANICA", "JURISDICCIÓN",
                               "LLAVE_CONCATENADA", "PROFESION 1", "ESTADO", "DIAS ORIGEN", "PERIODO", "días_origen", "periodo"]
                resultado_fase2 = resultado_fase2.drop(columns=[c for c in cols_quitar if c in resultado_fase2.columns], errors='ignore')

                resultado_fase2.columns = [normalizar_texto(c) for c in resultado_fase2.columns]
                resultado_fase2 = resultado_fase2.loc[:, ~resultado_fase2.columns.duplicated()].copy()

                orden_final = [
                    "RUN", "DV", "APELLIDO 1", "APELLIDO 2", "NOMBRES",
                    "CARGO ORIGEN", "GRADO ORIGEN", "SUELDO BASE ORIGEN",
                    "UNIDAD LABORAL ORIGEN", "COMUNA ORIGEN", "ASIG. ZONA COMUNA ORIGEN",
                    "MONTO ASIG. ZONA ORIGEN", "FOLIO", "ANO", "FECHA INICIO CERT. SUP.",
                    "FECHA TERMINO CERT. SUP.", "FECHA CERTIFICACION", "PLAZA",
                    "GRADO SUPLENCIA", "SUELDO BASE SUPLENCIA", "CALIDAD JURIDICA",
                    "UNIDAD LABORAL SUPLENCIA", "COMUNA SUPLENCIA", "ASIG. ZONA SUPLENCIA",
                    "MONTO ASIG. ZONA SUPLENCIA", "TELETRABAJO OTRA LOCALIDAD",
                    "COMUNA TT SUPLENCIA", "ZONA ASIG COMUNA", "FECHA INICIO SUPLENCIA",
                    "FECHA TERMINO SUPLENCIA", "DIAS", "DIA O MES",
                    "PROPORCIONAL ZONA SUPLENCIA", "PROPORCIONAL ZONA ORIGEN", "DIFERENCIA ASIG. ZONA"
                ]

                cols_presentes = [c for c in orden_final if c in resultado_fase2.columns]
                resultado_fase2 = resultado_fase2[cols_presentes]

                # Generar Excel en memoria para la descarga web
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    resultado_fase2.to_excel(writer, index=False, sheet_name='Matriz_Calculada')
                processed_data = output.getvalue()

                st.success(f"✅ ¡Proceso finalizado! Se calcularon {len(resultado_fase2)} registros exitosamente.")
                
                st.download_button(
                    label="📥 Descargar Matriz Final Excel",
                    data=processed_data,
                    file_name="Matriz_CAPJ_Final_Visual.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            except Exception as e:
                st.error(f"❌ Ocurrió un error al procesar las bases: {e}")
