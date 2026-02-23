import streamlit as st
import pandas as pd
import os
import sys
from dbfread import DBF
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import numpy as np
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
import datetime
import tempfile
from io import BytesIO
import base64

# --- FUNCIÓN AUXILIAR RUTAS ---
def resolver_ruta(ruta_relativa):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)

def obtener_entregas_excluidas(rutas_historicas):
    """
    Lee archivos de remates anteriores para identificar qué Entregas/Contratos 
    ya fueron procesados. Divide las entregas combinadas (ej: 'A / B').
    """
    excluidas = set()
    
    if not rutas_historicas:
        return excluidas
        
    if isinstance(rutas_historicas, str):
        rutas_historicas = [rutas_historicas]
        
    st.info(f"Analizando {len(rutas_historicas)} archivos históricos...")
    
    for ruta in rutas_historicas:
        try:
            df = pd.read_excel(ruta, header=6)
            
            col_obj = None
            if 'Entrega' in df.columns:
                col_obj = 'Entrega'
            elif 'Contrato' in df.columns:
                col_obj = 'Contrato'
            
            if col_obj:
                serie = df[col_obj].astype(str)
                
                for valor in serie:
                    partes = valor.split('/')
                    for parte in partes:
                        limpio = parte.strip().replace('.0', '')
                        if limpio:
                            excluidas.add(limpio)
                            
                st.success(f"{os.path.basename(ruta)}: Procesado correctamente.")
            else:
                st.warning(f"{os.path.basename(ruta)}: No se encontró columna 'Entrega' o 'Contrato'.")
                
        except Exception as e:
            st.error(f"Error leyendo histórico {os.path.basename(ruta)}: {e}")
            
    st.info(f"Total entregas únicas a excluir: {len(excluidas)}")
    return excluidas

def obtener_entregas_excluidas_hojas(rutas_historicas):
    """
    Lee los nombres de las HOJAS de los archivos históricos.
    En Celulosa, cada hoja es una Entrega ya procesada.
    """
    excluidas = set()
    
    if not rutas_historicas:
        return excluidas
        
    if isinstance(rutas_historicas, str):
        rutas_historicas = [rutas_historicas]
        
    st.info(f"Analizando pestañas de {len(rutas_historicas)} archivos históricos...")
    
    for ruta in rutas_historicas:
        try:
            wb = load_workbook(ruta, read_only=True, keep_links=False)
            
            for sheet_name in wb.sheetnames:
                limpio = sheet_name.strip()
                if limpio:
                    excluidas.add(limpio)
            
            st.success(f"{os.path.basename(ruta)}: Pestañas extraídas.")
            wb.close()
                
        except Exception as e:
            st.error(f"Error leyendo histórico {os.path.basename(ruta)}: {e}")
            
    st.info(f"Total entregas (hojas) a excluir: {len(excluidas)}")
    return excluidas

# ==========================================
#   NUEVA FUNCIÓN AUXILIAR DE FORMATO
# ==========================================
def agregar_cabecera_arauco(ws, datos):
    """
    Dibuja la cabecera estilo Arauco en la hoja activa (ws).
    """
    fuente_negrita = Font(bold=True, name='Calibri', size=11)
    borde_fino = Side(border_style="thin", color="000000")
    caja = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)
    alineacion_izq = Alignment(horizontal="left", vertical="center")
    alineacion_centro = Alignment(horizontal="center", vertical="center")

    # Fila 1
    ws['A1'] = "Nave";          ws['B1'] = datos['nave']
    ws['D1'] = "Exportador";    ws['E1'] = datos['exportador']
    
    # Fila 2
    ws['A2'] = "Destino";       ws['B2'] = datos['destino']
    ws['D2'] = "Embarcador";    ws['E2'] = datos['embarcador']
    
    # Fila 3
    ws['A3'] = "Reserva";       ws['B3'] = datos['reserva']
    ws['D3'] = "Carga";         ws['E3'] = datos['carga']
    
    # Fila 4
    ws['A4'] = "Contrato";      ws['B4'] = datos['contrato']
    ws['D4'] = "Tipo/Linea";    ws['E4'] = datos['linea']

    ws.merge_cells('B4:C4')

    for row in range(1, 5):
        celda_tit1 = ws.cell(row=row, column=1)
        celda_tit2 = ws.cell(row=row, column=4)
        
        celda_tit1.font = fuente_negrita
        celda_tit2.font = fuente_negrita
        
        celda_val1 = ws.cell(row=row, column=2)
        celda_val2 = ws.cell(row=row, column=5)
        
        for col in range(1, 6):
            celda = ws.cell(row=row, column=col)
            celda.border = caja
            celda.alignment = alineacion_izq

    ws['B4'].alignment = alineacion_centro

# ==========================================
#      LÓGICA DE MADERA (CORREGIDA)
# ==========================================
def procesar_madera(rutas):
    """
    1. Separa entregas compuestas (ej: "A / B" -> fila A, fila B).
    2. Agrupa por Producto para respetar pesos/volúmenes.
    3. Genera cabecera personalizada en Remate.
    """
    st.info("Iniciando procesamiento de Madera...")
    
    def separar_entregas_multiples(df, col_entrega):
        if col_entrega not in df.columns:
            return df
        
        df[col_entrega] = df[col_entrega].astype(str)
        df[col_entrega] = df[col_entrega].str.split('/')
        df = df.explode(col_entrega)
        df[col_entrega] = df[col_entrega].str.strip().str.replace(r'\.0$', '', regex=True)
        return df

    try:
        # 1. Cargar PROGRAMA
        programa = pd.read_excel(rutas['programa'])
        programa = separar_entregas_multiples(programa, "Entrega")
        
        if 'historico' in rutas and rutas['historico']:
            excluidas = obtener_entregas_excluidas(rutas['historico'])
            if excluidas:
                st.info(f"Filtrando {len(excluidas)} entregas históricas...")
                programa['Entrega_Str'] = programa['Entrega'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                programa = programa[~programa['Entrega_Str'].isin(excluidas)].copy()
                programa = programa.drop(columns=['Entrega_Str'])
                if programa.empty:
                    return False, "Todas las entregas del programa ya fueron procesadas en los históricos adjuntos.", []

        # 2. Cargar SALDOS
        if 'saldos' in rutas and rutas['saldos']:
            try:
                saldos = pd.read_excel(rutas['saldos'])
                saldos = separar_entregas_multiples(saldos, "Entrega")
                st.success("Archivo Saldos cargado y normalizado.")
            except Exception as e:
                st.warning(f"Error leyendo Saldos: {e}. Continuando sin él.")
                saldos = pd.DataFrame(columns=["Entrega", "Box Saldo"])
        else:
            saldos = pd.DataFrame(columns=["Entrega", "Box Saldo"])

        # 3. Cargar DESPACHO
        despacho = pd.read_excel(rutas['despacho'])
        
        # 4. Cargar DETALLE
        detalle = pd.read_excel(rutas['detalle'])
        
        # 5. Cargar INFORME
        consolidado = pd.read_excel(rutas['informe'])
        
        # 6. Cargar ZOOPP
        ruta_zoopp = rutas['zoopp']
        if ruta_zoopp.lower().endswith('.dbf'):
            st.info("Detectado archivo DBF. Cargando con dbfread...")
            try:
                table = DBF(ruta_zoopp, encoding='latin-1', char_decode_errors='ignore')
                zoopp = pd.DataFrame(iter(table))
                zoopp.columns = [c.lower() for c in zoopp.columns]
                mapeo_dbf = {
                    "loteof": "loteof,C,10", "vollote": "vollote,C,15",
                    "posped": "posped,N,6,0", "desmat": "desmat,C,40"
                }
                zoopp = zoopp.rename(columns=mapeo_dbf)
            except Exception as e:
                st.error(f"Error leyendo DBF: {e}")
                raise e
        else:
            zoopp = pd.read_excel(rutas['zoopp'])

        # --- RENOMBRAR COLUMNAS DESPACHO ---
        mapa_columnas_despacho = {
            "cor_ano": "COR_ANO,N,16,0", "cor_mov": "COR_MOV,N,16,0", "sigla": "SIGLA,C,4",
            "numero": "NUMERO,N,16,0", "dv": "DV,C,1", "cliente": "CLIENTE,C,20",
            "cod_puerto_destino": "COD_PUERTO,C,3", "puerto_destino": "PUERTO_DES,C,40",
            "cod_puerto_descarga": "COD_PUERTO1,C,3", "puerto_descarga": "PUERTO_DES1,C,40",
            "operacion": "OPERACION,N,16,0", "nave": "NAVE,C,40", "nro_reserva": "NRO_RESERV,C,15",
            "orden_embarque": "ORDEN_EMBA,C,255", "contrato": "CONTRATO,C,50", "deposito": "DEPOSITO,C,15",
            "fecha_despacho": "FECHA_DESP,D", "turno": "TURNO,N,16,0", "despachador": "DESPACHADO,C,100",
            "status": "STATUS,C,1", "sello": "SELLO,C,15", "peso": "PESO,N,16,0", "fechaanula": "FECHAANULA,D",
            "material": "MATERIAL,C,2", "des_fotos": "DES_FOTOS,C,25", "producto": "PRODUCTO,C,100",
            "isocode": "ISOCODE,C,4", "linea_cnt": "LINEA_CNT,C,15", "cant_paquetes": "CANT_PAQUE,N,16,0",
            "ite_volumen_d": "ITE_VOLUME,N,16,0", "terminal": "TERMINAL,C,100"
        }
        despacho = despacho.rename(columns=mapa_columnas_despacho)

        despacho = separar_entregas_multiples(despacho, "CONTRATO,C,50")

        # --- FILTRADO Y LÓGICA ---
        saldos['Box Saldo'] = pd.to_numeric(saldos['Box Saldo'], errors='coerce')
        
        entregas_con_saldo = saldos.loc[saldos["Box Saldo"] != 0, "Entrega"].unique()
        prog_filtrado = programa[
            (~programa["Entrega"].isin(entregas_con_saldo)) & 
            (programa["PRODINFO"].isin(["M.ASER.VERDE", "M.ASER. SECA", "M&B/SHOP","CLEARS","MDF MOLDURAS","MOLDURAS","BLANKS","SHOP","MOULDING&BETTER","M.PALL.SECA","M.PALL.VERDE","BASAS","AGLOMERADOS","MDF PANEL","PLYWOOD","TRUPAN","TABLERO","OSB","CHAPAS"]))
        ].copy()

        columnas_prog = ["Entrega", "Nave", "PRODINFO", "RESERVA", "DESTINO"]
        prog_filtrado = prog_filtrado[columnas_prog]
        
        try:
            nave_header = prog_filtrado["Nave"].dropna().iloc[0]
        except:
            nave_header = "SIN NAVE"

        # Construir Contenedor Despacho
        despacho['SIGLA,C,4'] = despacho['SIGLA,C,4'].astype(str).str.strip()
        despacho['NUMERO,N,16,0'] = despacho['NUMERO,N,16,0'].astype(str).str.strip()
        despacho['DV,C,1'] = despacho['DV,C,1'].astype(str).str.strip()

        def construir_contenedor(row):
            sigla = row['SIGLA,C,4']
            numero = row['NUMERO,N,16,0'].zfill(6)
            dv = row['DV,C,1']
            return f"{sigla}-{numero}-{dv}"

        def construir_NDESPACHO(row):
            cor1 = row['COR_ANO,N,16,0']
            cor2 = row['COR_MOV,N,16,0']
            return f"{cor1}-{cor2}"

        despacho['NDESPACHO'] = despacho.apply(construir_NDESPACHO, axis=1)
        despacho['CONTENEDOR'] = despacho.apply(construir_contenedor, axis=1)

        # Merge Programa - Despacho
        prog_filtrado = prog_filtrado.merge(
            despacho[["CONTENEDOR", "SELLO,C,15", "NDESPACHO", "CONTRATO,C,50", "PESO,N,16,0","NUMERO,N,16,0"]].rename(columns={"CONTRATO,C,50": "Entrega"}),
            on="Entrega", how="inner"
        )

        # Procesar Detalle
        mapa_detalle = {
            "ano": "ANO,N,16,0", "numero": "NUMERO,N,16,0", "fecha_consolidacion": "FECHA_CONS,D",
            "turno": "TURNO,N,16,0", "linea": "LINEA,C,15", "operacion": "OPERACION,N,16,0",
            "nave": "NAVE,C,47", "cliente": "CLIENTE,C,20", "embarcador": "EMBARCADOR,C,20",
            "reserva": "RESERVA,C,30", "pedido": "PEDIDO,C,50", "producto": "PRODUCTO,C,100",
            "agrupacion": "AGRUPACION,C,50", "fotos": "FOTOS,C,25", "pto_descarga": "PTO_DESCAR,C,40",
            "pto_final": "PTO_FINAL,C,40", "isocode": "ISOCODE,C,4", "medida": "MEDIDA,N,16,0",
            "sigla_cnt": "SIGLA_CNT,C,4", "contenedor": "CONTENEDOR,C,9", "tara": "TARA,N,16,0",
            "neto": "NETO,N,16,0", "fardos": "FARDOS,N,16,0", "unidades": "UNIDADES,N,16,0",
            "volumen": "VOLUMEN,N,17,4", "sello_linea": "SELLO_LINE,C,20", "sello_inspector": "SELLO_INSP,C,20",
            "dus": "DUS,C,255", "inf_gate": "INF_GATE,C,50", "embarcado": "EMBARCADO,C,1",
            "origen_carga": "ORIGEN_CAR,C,100", "deposito_origen": "DEPOSITO_O,C,20",
            "deposito_destino": "DEPOSITO_D,C,20", "fecha_packing": "FECHA_PACK,D",
            "cancelado": "CANCELADO,C,1", "observacion": "OBSERVACIO,C,255", "ind_aforo": "IND_AFORO,C,1",
            "restriccion_peso": "RESTRICCIO,N,16,0", "cod_nro_cnt": "COD_NRO_CN,N,16,0",
            "cod_dv_cnt": "COD_DV_CNT,C,1", "nro_despacho": "NRO_DESPAC,C,15", "peso_vgm": "PESO_VGM,N,17,2"
        }
        detalle = detalle.rename(columns=mapa_detalle)
        detalle['SELLO_LINE,C,20'] = detalle['SELLO_LINE,C,20'].astype(str).str.strip()
        detalle = detalle.drop_duplicates(subset=['SELLO_LINE,C,20'])

        prog_filtrado = prog_filtrado.merge(
            detalle[['SELLO_LINE,C,20', 'SELLO_INSP,C,20', 'DUS,C,255', 'RESTRICCIO,N,16,0','FECHA_CONS,D']],
            left_on='SELLO,C,15', right_on='SELLO_LINE,C,20', how='left'
        )
        prog_filtrado.drop(columns=['SELLO_LINE,C,20'], inplace=True)

        # Procesar Informe (Consolidado)
        mapa_consolidado = {
            "operacion": "OPERACION,N,16,0", "nave": "NAVE,C,40", "linea": "LINEA,C,15",
            "cliente": "CLIENTE,C,20", "proveedor": "PROVEEDOR,C,20", "cod_pto_destino": "COD_PTO_DE,C,3",
            "pto_destino": "PTO_DESTIN,C,40", "contrato": "CONTRATO,C,50", "sigla_cnt": "SIGLA_CNT,C,4",
            "nro_cnt": "NRO_CNT,N,16,0", "dv_cnt": "DV_CNT,C,1", "tara_cnt": "TARA_CNT,N,16,0",
            "orden_embarque": "ORDEN_EMBA,C,255", "sello": "SELLO,C,15", "orden_pedido": "ORDEN_PEDI,C,12",
            "codigo_barra": "CODIGO_BAR,C,50", "nro_paquete": "NRO_PAQUET,C,50", "material": "MATERIAL,C,50",
            "marca": "MARCA,C,20", "volumen": "VOLUMEN,N,17,4", "unid_volumen": "UNID_VOLUM,C,4",
            "espesor": "ESPESOR,N,17,4", "unid_espesor": "UNID_ESPES,C,4", "ancho": "ANCHO,N,17,4",
            "unid_ancho": "UNID_ANCHO,C,4", "largo": "LARGO,N,17,4", "unid_largo": "UNID_LARGO,C,4",
            "cant_piezas": "CANT_PIEZA,N,16,0", "peso": "PESO,N,17,4", "terminal": "TERMINAL,C,100",
            "bodega": "BODEGA,C,15", "fila": "FILA,C,4", "columna": "COLUMNA,C,4", "reserva": "RESERVA,C,30",
            "Cantidad_Pqts": "CANTIDAD_P,N,16,0","maxgross":"MAXGROSS"
        }
        consolidado = consolidado.rename(columns=mapa_consolidado)
        if "MAXGROSS" not in consolidado.columns:
            consolidado["MAXGROSS"] = 999999
        consolidado['NRO_CNT,N,16,0'] = consolidado['NRO_CNT,N,16,0'].astype(str).str.strip()
        consolidado['SIGLA_CNT,C,4'] = consolidado['SIGLA_CNT,C,4'].astype(str).str.strip()
        consolidado['DV_CNT,C,1'] = consolidado['DV_CNT,C,1'].astype(str).str.strip()

        def construir_contenedor_2(row):
            sigla = row['SIGLA_CNT,C,4']
            numero = row['NRO_CNT,N,16,0'].zfill(6)
            dv = row['DV_CNT,C,1']
            return f"{sigla}-{numero}-{dv}"

        consolidado['CONTENEDOR_2'] = consolidado.apply(construir_contenedor_2, axis=1)
        consolidado_filtrado = consolidado[
            consolidado["CONTENEDOR_2"].isin(prog_filtrado["CONTENEDOR"])
        ].copy()

        columnas_consolidado = [
            "CONTENEDOR_2", "TARA_CNT,N,16,0", "MATERIAL,C,50",
            "CODIGO_BAR,C,50", "ORDEN_PEDI,C,12",
            "PESO,N,17,4", "CONTRATO,C,50","MAXGROSS"
        ]
        consolidado_filtrado = consolidado_filtrado[columnas_consolidado]
        zoopp['loteof,C,10'] = zoopp['loteof,C,10'].astype(str).str.strip()
        zoopp['clase_merc'] = zoopp['clase_merc'].astype(str).str.strip()
        zoopp = zoopp.drop_duplicates(subset=['loteof,C,10'])
        consolidado_filtrado['CODIGO_BAR,C,50'] = (
            consolidado_filtrado['CODIGO_BAR,C,50']
            .astype(str)
            .str.strip()
        )

        consolidado_filtrado = consolidado_filtrado.merge(
            zoopp[['loteof,C,10', 'clase_merc']],
            left_on='CODIGO_BAR,C,50',
            right_on='loteof,C,10',
            how='left'
        )

        consolidado_filtrado.drop(columns=['loteof,C,10'], inplace=True)

        prog_filtrado['PRODINFO'] = prog_filtrado['PRODINFO'].astype(str).str.strip()

        resultado_final = consolidado_filtrado.merge(
            prog_filtrado,
            left_on=['CONTENEDOR_2', 'clase_merc'],
            right_on=['CONTENEDOR', 'PRODINFO'],
            how='left'
        )

        # Procesar ZOOPP
        zoopp['loteof,C,10'] = zoopp['loteof,C,10'].astype(str).str.strip()
        zoopp['vollote,C,15'] = zoopp['vollote,C,15'].astype(str).str.strip()
        resultado_final['CODIGO_BAR,C,50'] = resultado_final['CODIGO_BAR,C,50'].astype(str).str.strip()
        zoopp = zoopp.drop_duplicates(subset=['loteof,C,10'])

        resultado_filtrado_zoopp = resultado_final[resultado_final["CODIGO_BAR,C,50"].isin(zoopp["loteof,C,10"])].copy()
        resultado_filtrado_zoopp = resultado_filtrado_zoopp.merge(
            zoopp[['loteof,C,10', 'posped,N,6,0', 'desmat,C,40','vollote,C,15','clase_merc']],
            left_on='CODIGO_BAR,C,50', right_on='loteof,C,10', how='left'
        )

        resultado_filtrado_zoopp["PESO,N,16,0"] = pd.to_numeric(resultado_filtrado_zoopp["PESO,N,16,0"], errors='coerce')
        resultado_filtrado_zoopp["TARA_CNT,N,16,0"] = pd.to_numeric(resultado_filtrado_zoopp["TARA_CNT,N,16,0"], errors='coerce')
        resultado_filtrado_zoopp["VGM"] = (resultado_filtrado_zoopp["PESO,N,16,0"].fillna(0) + resultado_filtrado_zoopp["TARA_CNT,N,16,0"].fillna(0))
        resultado_filtrado_zoopp.drop(columns=["PESO,N,16,0"], inplace=True)

        resultado_filtrado_zoopp["vollote,C,15"] = (
            resultado_filtrado_zoopp["vollote,C,15"].astype(str).str.replace(",", ".", regex=False)
        )
        resultado_filtrado_zoopp["vollote,C,15"] = pd.to_numeric(resultado_filtrado_zoopp["vollote,C,15"], errors='coerce')
        resultado_filtrado_zoopp = resultado_filtrado_zoopp.dropna(subset=["vollote,C,15"])
        
        resultado_filtrado_zoopp = resultado_filtrado_zoopp.drop_duplicates(
            subset=["loteof,C,10", "Entrega", "CONTENEDOR_2"], 
            keep="first"
        )

        # =========================================================================
        # --- GENERAR REMATE 
        # =========================================================================
        remate = (
            resultado_filtrado_zoopp.groupby(["CONTENEDOR", "Entrega", "PRODINFO"]).agg({
                "RESERVA": "first",
                "DESTINO": "first",
                "SELLO,C,15": "first",
                "CODIGO_BAR,C,50": "count",       
                "PESO,N,17,4": "sum",             
                "vollote,C,15": "sum",
                "TARA_CNT,N,16,0": "first", 
                "MAXGROSS": "first"                         
            }).reset_index()
        )
        
        remate["PESO_BRUTO_TOTAL"] = remate["PESO,N,17,4"] + remate["TARA_CNT,N,16,0"]
        contenedores_con_sobrepeso = set(
            remate[remate["PESO_BRUTO_TOTAL"] >= remate["MAXGROSS"]]["CONTENEDOR"].unique()
        )        

        remate = remate.rename(columns={
            "CONTENEDOR": "Contenedor",
            "Entrega": "Entrega",
            "RESERVA": "Reserva",
            "DESTINO": "Pto Destino",
            "PRODINFO": "Clase de Producto",
            "SELLO,C,15": "Sello",
            "CODIGO_BAR,C,50": "Cantidad de Lotes",
            "PESO,N,17,4": "Peso Total (kg)",
            "vollote,C,15": "Volumen Total (m3)"
        })

        remate = remate[[
            "Entrega", "Reserva", "Pto Destino", "Clase de Producto",
            "Contenedor", "Sello", "Cantidad de Lotes", 
            "Peso Total (kg)", "Volumen Total (m3)"
        ]]
        remate = remate.sort_values(by=["Entrega", "Contenedor"])
        
        # Crear archivo en memoria
        output = BytesIO()
        remate.to_excel(output, index=False, startrow=6, engine='openpyxl')
        
        wb = load_workbook(output)
        ws = wb.active
        
        fecha_hoy = datetime.datetime.now().strftime("%d/%m/%Y")
        
        ws['A1'] = "INFORME  DE CONTENEDORES CONSOLIDADOS PARA EMBARQUE"
        ws['A3'] = "SAN VICENTE TERMINAL INTERNACIONAL"
        ws['A4'] = f"FECHA: {fecha_hoy}"
        ws['A5'] = f"NAVE: {nave_header}"
        
        bold_font = Font(bold=True)
        ws['A1'].font = bold_font
        ws['A3'].font = bold_font
        ws['A5'].font = bold_font
        ws['A1'].alignment = Alignment(horizontal="left")
        ws['A3'].alignment = Alignment(horizontal="left")
        ws['A4'].alignment = Alignment(horizontal="left")
        ws['A5'].alignment = Alignment(horizontal="left")

        columnas_a_fusionar = [1, 2, 3]  
        start_row = 8 
        
        if ws.max_row >= start_row:
            current_value = ws.cell(row=start_row, column=1).value
            merge_start = start_row

            for row in range(start_row + 1, ws.max_row + 1):
                value = ws.cell(row=row, column=1).value
                if value != current_value:
                    if merge_start != row - 1:
                        for col in columnas_a_fusionar:
                            ws.merge_cells(start_row=merge_start, start_column=col, end_row=row - 1, end_column=col)
                    current_value = value
                    merge_start = row
            if merge_start != ws.max_row:
                for col in columnas_a_fusionar:
                    ws.merge_cells(start_row=merge_start, start_column=col, end_row=ws.max_row, end_column=col)

        rojo_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        alignment_center = Alignment(horizontal="center", vertical="center")
        
        for row in ws.iter_rows(min_row=7):
            celda_contenedor = row[4]
            valor_contenedor = str(celda_contenedor.value).strip()
            
            if valor_contenedor in contenedores_con_sobrepeso:
                celda_contenedor.fill = rojo_fill
                
            for cell in row:
                cell.alignment = alignment_center
        
        remate_output = BytesIO()
        wb.save(remate_output)
        remate_output.seek(0)
        
        # --- GENERAR REMATE SAG ---
        remate_sag = (
            resultado_filtrado_zoopp.groupby(["CONTENEDOR", "Entrega"]).agg({
                "RESERVA": "first",
                "DESTINO": "first",
                "PRODINFO": "first",
                "SELLO,C,15": "first",
                "SELLO_INSP,C,20": "first",
                "CODIGO_BAR,C,50": "count",
                "PESO,N,17,4": "sum",
                "vollote,C,15": "sum"
            }).reset_index()
        )

        remate_sag = remate_sag.rename(columns={
            "CONTENEDOR": "Contenedor",
            "Entrega": "Entrega",
            "RESERVA": "Reserva",
            "DESTINO": "Pto Destino",
            "PRODINFO": "Clase de Producto",
            "SELLO,C,15": "Sello",
            "SELLO_INSP,C,20": "Sello Inspector",
            "CODIGO_BAR,C,50": "Cantidad de Lotes",
            "PESO,N,17,4": "Peso Total (kg)",
            "vollote,C,15": "Volumen Total (m3)"
        })

        remate_sag = remate_sag[
            ["Entrega", "Reserva", "Pto Destino", "Clase de Producto",
             "Contenedor", "Sello", "Sello Inspector",
             "Cantidad de Lotes", "Peso Total (kg)", "Volumen Total (m3)"]
        ]

        remate_sag = remate_sag.sort_values(by=["Entrega", "Contenedor"])
        remate_sag_output = BytesIO()
        remate_sag.to_excel(remate_sag_output, index=False, engine='openpyxl')
        remate_sag_output.seek(0)
       
        # =========================================================================
        # --- GENERAR PICKING ORIGINAL ---
        # =========================================================================
        resultado_filtrado_zoopp["PESO,N,17,4"] = pd.to_numeric(resultado_filtrado_zoopp["PESO,N,17,4"], errors='coerce')
        resultado_filtrado_zoopp["TARA_CNT,N,16,0"] = pd.to_numeric(resultado_filtrado_zoopp["TARA_CNT,N,16,0"], errors='coerce')
        resultado_filtrado_zoopp['FECHA_CONS,D'] = pd.to_datetime(resultado_filtrado_zoopp['FECHA_CONS,D'], dayfirst=True).dt.strftime('%d/%m/%Y')

        picking_cabecera = (
            resultado_filtrado_zoopp.groupby(["CONTENEDOR", "Entrega"]).agg({
                "SELLO,C,15": "first",
                "RESERVA": "first",
                "DUS,C,255": "first",
                "PESO,N,17,4": "sum",
                "TARA_CNT,N,16,0": "first",
                "FECHA_CONS,D": "first"
            }).reset_index()
        )

        picking_cabecera = picking_cabecera.rename(columns={
            "SELLO,C,15": "Sello", "RESERVA": "Reserva", "DUS,C,255": "DUS",
            "PESO,N,17,4": "Peso Bruto (kg)", "TARA_CNT,N,16,0": "Tara (kg)", "FECHA_CONS,D": "Fecha Contable","Entrega":"Entrega"
        })
        picking_cabecera["Peso Total (kg)"] = picking_cabecera["Peso Bruto (kg)"] + picking_cabecera["Tara (kg)"]
        
        vals_fijos = {
            "TPLST": "ZTPC", "Un Med Peso": "KG", "Un Med Tara": "KG", "Material Embalaje": "HC40",
            "Clase Med Transporte": "Z100", "Clave Flete": "0001", "Tipo Flete": "01",
            "Nombre Despachador": "A", "Rut Despachador": "1", "Nombre Chofer": "A",
            "RUT Chofer": "1", "Patente": "A", "Transportista": "50025", "Guia": "1", "Almacen Destino": "7004"
        }
        for k, v in vals_fijos.items():
            picking_cabecera[k] = v
            
        picking_cabecera["ID Cabecera"] = range(1, len(picking_cabecera) + 1)
        picking_cabecera = picking_cabecera.rename(columns={
            "CONTENEDOR": "ID Contenedor", "Tara (kg)": "Tara Contenedor", "Sello": "Sello Cont Nro",
            "Reserva": "Booking Nro", "Peso Bruto (kg)": "Peso Bruto Carga",
            "Peso Total (kg)": "Entrega Peso Total", "DUS": "DUS Nro"
        })
        
        cols_pick = ["ID Cabecera", "Entrega","Almacen Destino", "Fecha Contable", "Guia", "Transportista", "Patente",
                     "RUT Chofer", "Nombre Chofer", "Rut Despachador", "Nombre Despachador", "Tipo Flete",
                     "Clave Flete", "Clase Med Transporte", "Material Embalaje", "ID Contenedor", "Tara Contenedor",
                     "Un Med Tara", "Sello Cont Nro", "Booking Nro", "Peso Bruto Carga", "Un Med Peso", "DUS Nro",
                     "Entrega Peso Total", "TPLST"]
        picking_cabecera = picking_cabecera[cols_pick]

        # Tabla POSICION (Original)
        posicion = resultado_filtrado_zoopp.merge(
            picking_cabecera[['ID Cabecera', 'ID Contenedor', 'Entrega']],
            left_on=['CONTENEDOR', 'Entrega'],
            right_on=['ID Contenedor', 'Entrega'],
            how='left'
        )
        posicion['Cantidad'] = posicion.groupby(['ID Cabecera', 'CODIGO_BAR,C,50'])['CODIGO_BAR,C,50'].transform('count')
        posicion['ID Posicion'] = posicion.groupby('ID Cabecera')['CODIGO_BAR,C,50'].rank(method='dense').astype(int)
        
        posicion = posicion[['ID Cabecera', 'ID Posicion', 'CODIGO_BAR,C,50', 'Cantidad', 'PESO,N,17,4']]
        posicion = posicion.rename(columns={'CODIGO_BAR,C,50': 'Lote', 'PESO,N,17,4': 'Peso'})
        posicion['Unidad'] = "PQT"
        posicion = posicion[['ID Cabecera', 'ID Posicion', 'Lote', 'Cantidad', 'Unidad', 'Peso']]
        posicion = posicion.sort_values(by=["ID Cabecera", "ID Posicion"]).reset_index(drop=True)

        picking_output = BytesIO()
        with pd.ExcelWriter(picking_output, engine="openpyxl") as writer:
            picking_cabecera.to_excel(writer, sheet_name="Cabecera", index=False)
            posicion.to_excel(writer, sheet_name="Posicion", index=False)
        picking_output.seek(0)

        # =========================================================================
        # --- GENERAR PICKING NUEVO ---
        # =========================================================================
        picking_cabecera_nuevo = (
            resultado_filtrado_zoopp.groupby(["CONTENEDOR", "Entrega"]).agg({
                "SELLO,C,15": "first",
                "RESERVA": "first",
                "DUS,C,255": "first",
                "PESO,N,17,4": "sum",
                "TARA_CNT,N,16,0": "first",
                "FECHA_CONS,D": "first"
            }).reset_index()
        )

        picking_cabecera_nuevo = picking_cabecera_nuevo.rename(columns={
            "CONTENEDOR": "ID Contenedor",
            "SELLO,C,15": "Sello Cont Nro", 
            "RESERVA": "Booking Nro", 
            "DUS,C,255": "DUS Nro",
            "PESO,N,17,4": "Peso Bruto Carga", 
            "TARA_CNT,N,16,0": "Tara Contenedor", 
            "FECHA_CONS,D": "Fecha Contable"
        })
        
        picking_cabecera_nuevo["Peso Total"] = picking_cabecera_nuevo["Peso Bruto Carga"] + picking_cabecera_nuevo["Tara Contenedor"]
        picking_cabecera_nuevo["ID Cabecera"] = range(1, len(picking_cabecera_nuevo) + 1)
        
        vals_fijos_nuevo = {
            "Centro Origen": "TD06",
            "Almacen Origen": "0100",
            "Centro Destino": "TD06",
            "Almacen Destino": "7004",
            "Guia": "1",
            "Transportista": "50025",
            "Patente": "A",
            "RUT Chofer": "1",
            "Nombre Chofer": "A",
            "Rut Despachador": "1",
            "Nombre Despachador": "A",
            "Tipo Flete": "01",
            "Clave Flete": "0001",
            "Clase Med Transporte": "Z100",
            "Material Embalaje": "HC40",
            "Un Medida Tara": "KG",
            "Un Med Peso": "KG",
            "TPLST": "ZTPC"
        }
        for k, v in vals_fijos_nuevo.items():
            picking_cabecera_nuevo[k] = v
            
        cols_pick_nuevo = [
            "ID Cabecera", "Centro Origen", "Almacen Origen", "Centro Destino", 
            "Almacen Destino", "Fecha Contable", "Guia", "Transportista", 
            "Patente", "RUT Chofer", "Nombre Chofer", "Rut Despachador", 
            "Nombre Despachador", "Tipo Flete", "Clave Flete", "Clase Med Transporte", 
            "Material Embalaje", "ID Contenedor", "Tara Contenedor", "Un Medida Tara", 
            "Sello Cont Nro", "Booking Nro", "Peso Bruto Carga", "Un Med Peso", 
            "DUS Nro", "Entrega", "Peso Total", "TPLST"
        ]
        picking_cabecera_nuevo = picking_cabecera_nuevo[cols_pick_nuevo]

        # Tabla POSICION (Nuevo)
        posicion_nuevo = resultado_filtrado_zoopp.merge(
            picking_cabecera_nuevo[['ID Cabecera', 'ID Contenedor', 'Entrega']],
            left_on=['CONTENEDOR', 'Entrega'],
            right_on=['ID Contenedor', 'Entrega'],
            how='left'
        )
        posicion_nuevo['Cantidad'] = posicion_nuevo.groupby(['ID Cabecera', 'CODIGO_BAR,C,50'])['CODIGO_BAR,C,50'].transform('count')
        posicion_nuevo['ID Posicion'] = posicion_nuevo.groupby('ID Cabecera')['CODIGO_BAR,C,50'].rank(method='dense').astype(int)
        
        posicion_nuevo = posicion_nuevo.rename(columns={
            'CODIGO_BAR,C,50': 'Lote', 
            'PESO,N,17,4': 'Peso', 
            'ID Contenedor': 'BOX' 
        })
        posicion_nuevo['Unidad'] = "PQT"
        
        posicion_nuevo = posicion_nuevo[['ID Cabecera', 'ID Posicion', 'Lote', 'Cantidad', 'Unidad', 'Peso', 'BOX']]
        posicion_nuevo = posicion_nuevo.sort_values(by=["ID Cabecera", "ID Posicion"]).reset_index(drop=True)

        picking_nuevo_output = BytesIO()
        with pd.ExcelWriter(picking_nuevo_output, engine="openpyxl") as writer:
            picking_cabecera_nuevo.to_excel(writer, sheet_name="Cabecera", index=False)
            posicion_nuevo.to_excel(writer, sheet_name="Posicion", index=False)
        picking_nuevo_output.seek(0)

        # RETORNAMOS LOS 4 ARCHIVOS EN EL ARREGLO FINAL
        return True, "Proceso completado exitosamente", [
            ("RemateMadera.xlsx", remate_output),
            ("RemateMaderaSAG.xlsx", remate_sag_output),
            ("Picking.xlsx", picking_output),
            ("Picking_Nuevo.xlsx", picking_nuevo_output)
        ]

    except Exception as e:
        st.error(f"Error en procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, str(e), []

# ==========================================
#      LÓGICA DE Celulosa BKP EKP UKP
# ==========================================
def procesar_celulosa_cb(rutas):
    st.info("Iniciando procesamiento de Celulosa BKP EKP UKP...")
    try:
        programa = pd.read_excel(rutas['programa'])
        tools_celulosa = pd.read_excel(rutas['tools'])

        if 'saldos' in rutas and rutas['saldos']:
            try:
                saldos = pd.read_excel(rutas['saldos'])
                st.success("Saldos cargado.")
            except:
                saldos = pd.DataFrame(columns=["Entrega", "Box Saldo"])
        else:
            saldos = pd.DataFrame(columns=["Entrega", "Box Saldo"])

        saldos['Entrega'] = saldos['Entrega'].astype(str).str.strip()
        programa['Entrega'] = programa['Entrega'].astype(str).str.strip()
        saldos['Box Saldo'] = pd.to_numeric(saldos['Box Saldo'], errors='coerce')
        
        if 'historico' in rutas and rutas['historico']:
            excluidas = obtener_entregas_excluidas_hojas(rutas['historico'])
            if excluidas:
                st.info(f"Filtrando contra {len(excluidas)} entregas históricas (Pestañas)...")
                programa = programa[~programa['Entrega'].isin(excluidas)].copy()
                
                if programa.empty:
                    return False, "Todas las entregas del programa ya existen como hojas en los históricos adjuntos.", []

        entregas_con_saldo = saldos.loc[saldos["Box Saldo"] != 0, "Entrega"].unique()
        
        prog_filtrado = programa[
            (~programa["Entrega"].isin(entregas_con_saldo)) & 
            (programa["PRODINFO"].isin(["CEL BKP", "CEL UKP", "CEL EKP"]))
        ].copy()

        def obtener_linea(nav):
            nav = str(nav).upper().strip()
            if "MSC" in nav: return "MSC"
            if "ONEY" in nav: return "ONE"
            if "HLL" in nav: return "HAPAG LLOYD"
            if "MAERSK" in nav or "ML" in nav: return "MAERSK"
            return nav

        prog_filtrado['NAV_CLEAN'] = prog_filtrado['NAV'].apply(obtener_linea)
        
        metadata_dict = prog_filtrado.set_index('Entrega')[
            ['Nave', 'DESTINO', 'RESERVA', 'PRODINFO', 'NAV_CLEAN']
        ].to_dict('index')

        tools_celulosa['Contrato'] = tools_celulosa['Contrato'].astype(str)
        entregas_validas = prog_filtrado["Entrega"].unique()
        tools_filtrado = tools_celulosa[
            tools_celulosa["Contrato"].isin(entregas_validas)
        ]

        df = tools_filtrado.copy()
        df["Contenedor"] = df["Contenedor"].astype(str).str.strip()
        df["Expedicion"] = df["Expedicion"].astype(str).str.strip()
        
        def normalizar_box(contenedor):
            partes = contenedor.split('-')
            if len(partes) == 3:
                parte_media_normalizada = partes[1].zfill(6)
                return f"{partes[0]}-{parte_media_normalizada}-{partes[2]}"
            return contenedor

        df["BOX"] = df["Contenedor"].apply(normalizar_box)

        df["TARA"] = df["Tara"]
        df["LOTE"] = df["Expedicion"]
        df["BULTOS"] = df["Cantidad"]
        df["UNI"] = df["BULTOS"] / 8
        df["SELLO"] = df["Sello_linea"]
        df["RESERVA"] = df["Reserva"]
        df["DUS"] = df["Orden_Embarque"]
        df["MAX"] = df["Max_Gross"]

        df_agrupado = (
            df.groupby(["Contrato", "BOX", "LOTE"], as_index=False)
              .agg({
                  "TARA": "first",
                  "BULTOS": "sum",
                  "SELLO": "first",
                  "RESERVA": "first",
                  "DUS": "first",
                  "MAX": "first"
              })
        )
        
        df_agrupado["UNI"] = df_agrupado["BULTOS"] / 8
        columnas_finales = ["BOX", "TARA", "BULTOS", "UNI", "LOTE", "SELLO", "RESERVA", "DUS", "MAX"]
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for contrato, data in df_agrupado.groupby("Contrato"):
                data_limpia = data[columnas_finales]
                
                data_limpia.to_excel(writer, sheet_name=str(contrato), index=False, startrow=5)
                
                ws = writer.sheets[str(contrato)]
                meta = metadata_dict.get(str(contrato), {})
                
                datos_cabecera = {
                    'nave': meta.get('Nave', ''),
                    'destino': meta.get('DESTINO', ''),
                    'reserva': meta.get('RESERVA', ''),
                    'contrato': str(contrato),
                    'exportador': "ARAUCO",
                    'embarcador': "CELULOSA ARAUCO",
                    'carga': meta.get('PRODINFO', ''),
                    'linea': meta.get('NAV_CLEAN', '')
                }
                agregar_cabecera_arauco(ws, datos_cabecera)

        wb = load_workbook(output)

        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            
            header_row_idx = 6
            idx_box = 1
            for cell in ws[header_row_idx]:
                if cell.value == "BOX":
                    idx_box = cell.col_idx
                    break
            
            max_row = ws.max_row
            start = header_row_idx + 1
            
            while start <= max_row:
                valor = ws.cell(row=start, column=idx_box).value
                end = start
                while end + 1 <= max_row and ws.cell(row=end + 1, column=idx_box).value == valor:
                    end += 1
                
                if end > start:
                    ws.merge_cells(start_row=start, start_column=idx_box, end_row=end, end_column=idx_box)
                    ws.cell(row=start, column=idx_box).alignment = Alignment(vertical="center")
                
                start = end + 1

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        return True, "Archivo generado correctamente", [("CelulosaBKPEKPUKP.xlsx", final_output)]

    except Exception as e:
        st.error(f"Error en procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, str(e), []

# ==========================================
#      LÓGICA DE CELULOSA DP
# ==========================================
def procesar_celulosa_sb(rutas):
    st.info("Iniciando procesamiento de Celulosa DP...")
    try:
        programa = pd.read_excel(rutas['programa'])
        informe = pd.read_excel(rutas['informe'])

        if 'saldos' in rutas and rutas['saldos']:
            try:
                saldos = pd.read_excel(rutas['saldos'])
            except:
                saldos = pd.DataFrame(columns=["Entrega", "Box Saldo"])
        else:
            saldos = pd.DataFrame(columns=["Entrega", "Box Saldo"])

        saldos['Entrega'] = saldos['Entrega'].astype(str).str.strip()
        programa['Entrega'] = programa['Entrega'].astype(str).str.strip()
        saldos['Box Saldo'] = pd.to_numeric(saldos['Box Saldo'], errors='coerce')
        
        if 'historico' in rutas and rutas['historico']:
            excluidas = obtener_entregas_excluidas_hojas(rutas['historico'])
            if excluidas:
                st.info(f"Filtrando contra {len(excluidas)} entregas históricas (Pestañas)...")
                programa = programa[~programa['Entrega'].isin(excluidas)].copy()
                
                if programa.empty:
                    return False, "Todas las entregas del programa ya existen como hojas en los históricos adjuntos.", []

        entregas_con_saldo = saldos.loc[saldos["Box Saldo"] != 0, "Entrega"].unique()
        
        prog_filtrado = programa[
            (~programa["Entrega"].isin(entregas_con_saldo)) & 
            (programa["PRODINFO"].isin(["CEL DP"]))
        ].copy()
        
        def obtener_linea(nav):
            nav = str(nav).upper().strip()
            if "MSC" in nav: return "MSC"
            if "ONEY" in nav: return "ONE"
            if "HLL" in nav: return "HAPAG LLOYD"
            if "MAERSK" in nav or "ML" in nav: return "MAERSK"
            return nav

        prog_filtrado['NAV_CLEAN'] = prog_filtrado['NAV'].apply(obtener_linea)
        
        metadata_dict = prog_filtrado.set_index('Entrega')[
            ['Nave', 'DESTINO', 'RESERVA', 'PRODINFO', 'NAV_CLEAN']
        ].to_dict('index')

        entregas_validas = prog_filtrado["Entrega"].unique()
        
        if "contrato" in informe.columns:
            informe["contrato"] = informe["contrato"].astype(str).str.strip()
            informe = informe[informe["contrato"].isin(entregas_validas)]
        else:
            return False, "El archivo Informe no tiene la columna 'contrato'.", []

        informe['nro_cnt'] = informe['nro_cnt'].astype(str).str.strip()
        informe['sigla_cnt'] = informe['sigla_cnt'].astype(str).str.strip()
        informe['dv_cnt'] = informe['dv_cnt'].astype(str).str.strip()

        def construir_contenedor_2(row):
            sigla = row['sigla_cnt']
            numero = row['nro_cnt'].zfill(6)
            dv = row['dv_cnt']
            return f"{sigla}-{numero}-{dv}"

        informe['CONTENEDOR_2'] = informe.apply(construir_contenedor_2, axis=1)

        df = informe.rename(columns={
            "CONTENEDOR_2": "BOX",
            "tara_cnt": "TARA",
            "marca": "LOTE",
            "sello": "SELLO",
            "orden_embarque": "RESERVA",
            "reserva": "DUS",
            "maxgross": "MAX"
        })

        df = df[df["SELLO"].notna() & (df["SELLO"].astype(str).str.strip() != "")]

        agrupado = df.groupby(["BOX", "LOTE"]).agg({
            "TARA": "first",
            "SELLO": "first",
            "RESERVA": "first",
            "DUS": "first",
            "contrato": "first",
            "MAX":"first",
            "BOX": "count"
        }).rename(columns={"BOX": "UNI"})

        agrupado["BULTOS"] = agrupado["UNI"] * 8

        agrupado = agrupado.reset_index()[[
            "BOX", "TARA", "BULTOS", "UNI", "LOTE",
            "SELLO", "RESERVA", "DUS", "MAX", "contrato"
        ]]

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for contrato, data in agrupado.groupby("contrato"):
                hoja = str(contrato)
                data_limpia = data.drop(columns=["contrato"], inplace=False)
                
                data_limpia.to_excel(writer, sheet_name=hoja, index=False, startrow=5)

                ws = writer.sheets[hoja]
                meta = metadata_dict.get(hoja, {})
                
                datos_cabecera = {
                    'nave': meta.get('Nave', ''),
                    'destino': meta.get('DESTINO', ''),
                    'reserva': meta.get('RESERVA', ''),
                    'contrato': hoja,
                    'exportador': "ARAUCO",
                    'embarcador': "CELULOSA ARAUCO",
                    'carga': meta.get('PRODINFO', ''),
                    'linea': meta.get('NAV_CLEAN', '')
                }
                
                agregar_cabecera_arauco(ws, datos_cabecera)
        
        wb = load_workbook(output)

        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            
            header_row_idx = 6
            idx_box = 1
            for cell in ws[header_row_idx]:
                if cell.value == "BOX":
                    idx_box = cell.col_idx
                    break
            
            max_row = ws.max_row
            start = header_row_idx + 1
            
            while start <= max_row:
                valor = ws.cell(row=start, column=idx_box).value
                end = start
                while end + 1 <= max_row and ws.cell(row=end + 1, column=idx_box).value == valor:
                    end += 1
                
                if end > start:
                    ws.merge_cells(start_row=start, start_column=idx_box, end_row=end, end_column=idx_box)
                    ws.cell(row=start, column=idx_box).alignment = Alignment(vertical="center")
                
                start = end + 1

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        return True, "Archivo generado correctamente", [("RemateCelulosaDP.xlsx", final_output)]

    except Exception as e:
        st.error(f"Error en procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, str(e), []

# ==========================================
#      LÓGICA DE SAG 
# ==========================================
def procesar_sag(rutas):
    st.info("Iniciando procesamiento de SAG...")
    try:
        path_remate = rutas['remate']
        remate = pd.read_excel(path_remate)
        
        rutas_sif = rutas['sag']
        
        if isinstance(rutas_sif, str):
            rutas_sif = [rutas_sif]
            
        lista_sifs = []
        st.info(f"Cargando {len(rutas_sif)} archivos SIF...")
        
        for ruta in rutas_sif:
            try:
                df_temp = pd.read_excel(ruta, sheet_name="detalle")
                lista_sifs.append(df_temp)
            except Exception as e:
                st.error(f"Error cargando {ruta}: {e}")
        
        if not lista_sifs:
            return False, "No se pudo cargar ningún archivo SIF válido.", []
            
        SAG = pd.concat(lista_sifs, ignore_index=True)
        
        path_picking = rutas['picking']
        
        if not os.path.exists(path_picking):
             return False, f"No se encontró el archivo Picking: {path_picking}", []

        picking_pos = pd.read_excel(path_picking, sheet_name="Posicion")
        picking_cab = pd.read_excel(path_picking, sheet_name="Cabecera")

        # Normalizar columnas claves
        if "Codigo_Barra" in SAG.columns:
            SAG["Codigo_Barra"] = SAG["Codigo_Barra"].astype(str).str.strip()
        else:
            return False, "Los archivos SIF no tienen la columna 'Codigo_Barra'."
            
        picking_pos["Lote"] = picking_pos["Lote"].astype(str).str.strip()
        
        if "SIF" in SAG.columns:
            # 1. Crear una columna numérica temporal para ordenar correctamente
            SAG['SIF_num'] = pd.to_numeric(SAG['SIF'], errors='coerce')
            
            # 2. Ordenar por Codigo_Barra y SIF_num (descendente para que el mayor quede arriba)
            SAG = SAG.sort_values(by=['Codigo_Barra', 'SIF_num'], ascending=[True, False])
            
            # 3. Eliminar duplicados de lote, conservando el primero (que ahora es el SIF mayor)
            SAG = SAG.drop_duplicates(subset=['Codigo_Barra'], keep='first')
            
            # 4. Normalizar SIF a texto (tu lógica original)
            SAG['SIF'] = (
                SAG['SIF']
                .astype(str)
                .str.strip()
                .str.replace(r'\.0$', '', regex=True)
            )
            
            # Opcional: Eliminar la columna temporal si ya no la necesitas
            SAG = SAG.drop(columns=['SIF_num'])
        else:
            return False, "Los archivos SIF no tienen la columna 'SIF'."

        # Merge Picking Posicion con SIF
        picking_pos = picking_pos.merge(
            SAG[["Codigo_Barra", "SIF"]],
            how="left",
            left_on="Lote",
            right_on="Codigo_Barra"
        )
        
        if "Codigo_Barra" in picking_pos.columns:
            picking_pos = picking_pos.drop(columns=["Codigo_Barra"])

        picking_pos["SIF"] = picking_pos["SIF"].astype(str).str.strip()
        picking_pos["Lote"] = picking_pos["Lote"].astype(str).str.strip()

        sif_por_cabecera = (
            picking_pos[["ID Cabecera", "SIF"]]
            .dropna()
            .drop_duplicates()
        )

        metricas = (
            picking_pos.groupby(["ID Cabecera", "SIF"])
            .agg({
                "Lote": "count",        # Cantidad de Lotes
                "Peso": "sum"           # Peso Total
            })
            .reset_index()
            .rename(columns={
                "Lote": "Cantidad de Lotes",
                "Peso": "Peso Total"
            })
        )

        picking_cab = picking_cab.merge(
            sif_por_cabecera,
            on="ID Cabecera",
            how="left"
        )

        picking_cab = picking_cab.merge(
            metricas,
            on=["ID Cabecera", "SIF"],
            how="left"
        )

        picking_cab["ID Contenedor"] = picking_cab["ID Contenedor"].astype(str).str.strip()
        remate["Contenedor"] = remate["Contenedor"].astype(str).str.strip()

        remate = remate.merge(
            picking_cab[[
                "ID Contenedor",
                "SIF",
                "Cantidad de Lotes",
                "Peso Total"
            ]],
            left_on="Contenedor",
            right_on="ID Contenedor",
            how="left"
        )

        remate = remate.drop(columns=["ID Contenedor"], errors="ignore")
        
        if "Cantidad de Lotes_x" in remate.columns:
            remate = remate.drop(columns=["Cantidad de Lotes_x"])

        if "Cantidad de Lotes_y" in remate.columns:
            remate = remate.rename(columns={"Cantidad de Lotes_y": "Cantidad de Lotes"})

        remate = remate.rename(columns={"Peso Total": "Peso Lote"})

        output = BytesIO()
        remate.to_excel(output, index=False, engine='openpyxl')
        
        wb = load_workbook(output)
        ws = wb.active

        columnas_merge = [5, 6, 7, 8, 9]

        fila_inicio = 6
        while fila_inicio <= ws.max_row:
            fila_fin = fila_inicio

            while (
                fila_fin + 1 <= ws.max_row and
                all(
                    ws.cell(row=fila_inicio, column=col).value ==
                    ws.cell(row=fila_fin + 1, column=col).value
                    for col in columnas_merge
                )
            ):
                fila_fin += 1

            if fila_fin > fila_inicio:
                for col in columnas_merge:
                    ws.merge_cells(
                        start_row=fila_inicio,
                        start_column=col,
                        end_row=fila_fin,
                        end_column=col
                    )

            fila_inicio = fila_fin + 1

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        return True, "Archivo generado correctamente", [("RemateSIF.xlsx", final_output)]

    except Exception as e:
        st.error(f"Error en procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, str(e), []

# ==========================================
#      LÓGICA CMPC CELULOSA
# ==========================================
def procesar_cmpc_celulosa(rutas):
    st.info("Iniciando procesamiento CMPC Celulosa...")
    try:
        remate = pd.read_excel(rutas['remate'])
        tools = pd.read_excel(rutas['tools'])

        if 'producto' in remate.columns:
            remate = remate[remate['producto'] != "PAPEL KRAFT"]
        
        if "sello_linea" in remate.columns:
            remate["sello_linea_clean"] = (
                remate["sello_linea"]
                .astype(str)
                .str.replace("-", "", regex=False)
                .str.strip()
            )
        else:
             return False, "Columna 'sello_linea' no encontrada en Remate.", []

        if "Sello_linea" in tools.columns:
            tools["Sello_linea_clean"] = (
                tools["Sello_linea"]
                .astype(str)
                .str.replace("-", "", regex=False)
                .str.strip()
            )
        else:
            return False, "Columna 'Sello_linea' no encontrada en Tools.", []

        sellos_validos = set(remate["sello_linea_clean"])
        tools_filtrado = tools[tools["Sello_linea_clean"].isin(sellos_validos)].copy()

        df = tools_filtrado.merge(
            remate,
            left_on="Sello_linea_clean",
            right_on="sello_linea_clean",
            how="left",
            suffixes=("_tools", "_remate")
        )

        consolidado = pd.DataFrame()
        consolidado["Etiqueta"] = df["Expedicion"]
        consolidado["Contenedor"] = df["Contenedor"]
        consolidado["Sello"] = df["Sello_linea_clean"]
        consolidado["Tara"] = pd.to_numeric(df["Tara"], errors="coerce")
        consolidado["Tipo Cont."] = df["Tipo_Contenedor"]
        consolidado["Dimension"] = df["medida"]
        consolidado["Naviera"] = df["linea"]
        consolidado["Reserva"] = df["reserva"]
        consolidado["Dus"] = df["dus"]
        consolidado["agencia"] = df["aga"]
        consolidado["Bodega"] = ""
        consolidado["ubicación"] = ""
        consolidado["Directo"] = "N"
        consolidado["Destino"] = df["Pto_Destino"].astype(str).str.split(",", n=1).str[0]
        consolidado["Fardos"] = pd.to_numeric(df["Cantidad"], errors="coerce")
        consolidado["Pedido"] = df["Contrato"].astype(str).str.split("-", n=1).str[0]
        consolidado["fecha dus"] = (
            pd.to_datetime(df["fecha_aceptacion"], format="%d/%m/%Y %H:%M", errors="coerce")
            .dt.strftime("%d/%m/%Y")
        )
        consolidado["UNIT"] = consolidado["Fardos"] / 8
        consolidado["PLANTA"] = (
            df["producto"]
            .astype(str)
            .str.upper()
            .str.replace("CELULOSA ", "", regex=False)
            .str.strip()
        )
        consolidado["Peso neto"] = 0.25175 * consolidado["Fardos"]
        consolidado["Peso bruto"] = 0.25413 * consolidado["Fardos"]
        consolidado["Peso Total"] = 24396 + consolidado["Tara"]

        def calcular_volumen(planta, fardos):
            if pd.isna(fardos):
                return np.nan
            planta = str(planta).upper().strip()

            if "STA" in planta or "FÉ" in planta:
                return fardos * 0.254
            elif "LAJA" in planta:
                return fardos * 0.2502
            elif "PACIFICO" in planta:
                return fardos * 0.2618
            return np.nan

        consolidado["Volumen"] = [
            calcular_volumen(p, f)
            for p, f in zip(consolidado["PLANTA"], consolidado["Fardos"])
        ]

        consolidado["Marca"] = consolidado["Etiqueta"].astype(str) + "/" + consolidado["PLANTA"].astype(str)

        output = BytesIO()
        consolidado.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        
        return True, "Archivo generado", [("CMPC_Celulosa_Consolidado.xlsx", output)]

    except Exception as e:
        st.error(f"Error en procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, str(e), []

# ==========================================
#      LÓGICA CMPC MADERA (FINAL - NOTA POR CONTENEDOR)
# ==========================================
def procesar_cmpc_madera(rutas):
    st.info("Iniciando procesamiento CMPC Madera...")
    try:
        remate = pd.read_excel(rutas['remate'])
        tools = pd.read_excel(rutas['informe'])

        remate['sigla_cnt'] = remate['sigla_cnt'].astype(str).str.strip()
        remate['nro_cnt'] = remate['nro_cnt'].astype(str).str.strip()
        remate['dv_cnt'] = remate['dv_cnt'].astype(str).str.strip()

        def construir_contenedor_rem(row):
            sigla = str(row['sigla_cnt']).strip()
            val_num = str(row['nro_cnt'])
            if '.' in val_num:
                numero = val_num.split('.')[0].strip()
            else:
                numero = val_num.strip()
            dv = str(row['dv_cnt']).strip()
            numero = numero.zfill(6)
            contenedor = f"{sigla}-{numero}-{dv}"
            return contenedor

        remate['CONTENEDORREM'] = remate.apply(construir_contenedor_rem, axis=1)

        if "Sello_linea" in tools.columns:
            tools["Sello_linea_clean"] = (
                tools["Sello_linea"]
                .astype(str)
                .str.replace("-", "", regex=False)
                .str.strip()
            )

        cols_tools_necesarias = ['Cnt_Sigla', 'Cnt_Nro', 'Cnt_DV']
        for col in cols_tools_necesarias:
            if col not in tools.columns:
                return False, f"El archivo Informe (Tools) no tiene la columna '{col}'", []

        tools['Cnt_Sigla'] = tools['Cnt_Sigla'].astype(str).str.strip()
        tools['Cnt_Nro'] = tools['Cnt_Nro'].astype(str).str.strip()
        tools['Cnt_DV'] = tools['Cnt_DV'].astype(str).str.strip()

        def construir_contenedor_tools(row):
            sigla = str(row['Cnt_Sigla']).strip()
            val_num = str(row['Cnt_Nro'])
            if '.' in val_num:
                numero = val_num.split('.')[0].strip()
            else:
                numero = val_num.strip()
            dv = str(row['Cnt_DV']).strip()
            numero = numero.zfill(6)
            contenedor = f"{sigla}-{numero}-{dv}"
            return contenedor

        tools['CONTENEDORINF'] = tools.apply(construir_contenedor_tools, axis=1)

        mensajes_exito = []
        archivos_output = []

        # SUB-PROCESO 1: MADERA SECA
        remate_seca = remate[remate["producto"] == "MADERA SECA"].copy()
        
        if not remate_seca.empty:
            try:
                remate_seca['Desc_Carga_Calc'] = remate_seca['cant_piezas'].astype(str) + " PIECES, CHILEAN RADIATA PINE"
                
                contenedores_unicos_s = remate_seca['CONTENEDORREM'].unique()
                mapa_nota_s = {cnt: i+1 for i, cnt in enumerate(contenedores_unicos_s)}

                df_remate_extra_seca = pd.DataFrame({
                    "Nota": remate_seca['CONTENEDORREM'].map(mapa_nota_s),
                    "Venta": remate_seca["pedido"],
                    "Reserva": remate_seca["reserva"],
                    "Contenedor": remate_seca["CONTENEDORREM"],
                    "Sello Naviera (Carrier Seal)": remate_seca["sello_linea"],
                    "Descripción de la Carga": remate_seca["Desc_Carga_Calc"],
                    "N° de Pqts.": remate_seca["cant_paquetes"],
                    "Tara del Contenedor": remate_seca["tara"],
                    "Volumen Bruto de la Carga": remate_seca["volumen"],
                    "Peso Bruto de la Carga (documental)": remate_seca["neto"],
                    "Volumen Bruto del Contenedor": remate_seca["volumen"], 
                    "Comentarios del Contenedor": remate_seca["pto_final"]
                })
                
                output_seca_remate = BytesIO()
                df_remate_extra_seca.to_excel(output_seca_remate, index=False, engine='openpyxl')
                output_seca_remate.seek(0)
                archivos_output.append(("Remate_CMPC_Madera_Seca.xlsx", output_seca_remate))
                
            except Exception as e:
                st.warning(f"Error generando Remate Extra Seca: {e}")

            # CONSOLIDADO SECA
            tools_filtrado = tools[tools['CONTENEDORINF'].isin(remate_seca['CONTENEDORREM'])]
            remate_matched = remate_seca.set_index("CONTENEDORREM")
            tools_matched = tools_filtrado.set_index("CONTENEDORINF")
            
            df = tools_matched.join(remate_matched, how="left", rsuffix="_rem")
            
            if not df.empty:
                df[['contrato', 'item']] = df['Orden_Pedido'].astype(str).str.split('-', n=1, expand=True)
                df['fecha_dus'] = pd.to_datetime(df['fecha_aceptacion'], errors='coerce').dt.strftime('%d/%m/%Y')

                df_consolidado = pd.DataFrame({
                    "Npaquete": df["Nro_Paquete"],
                    "Contenedor": df.index,
                    "Sello": df["sello_linea"],
                    "Tara": df["tara"],
                    "Dimension": df["medida"],
                    "Tipo Cont.": df["tipo"],
                    "Directo": "N",
                    "Destino": df["pto_final"],
                    "Paquetes": "1",
                    "contrato": df["contrato"],
                    "item": df["item"],
                    "Naviera": df["linea"],
                    "Bodega": "",
                    "ubicación": "",
                    "Reserva": df["reserva"],
                    "Dus": df["dus"],
                    "fecha dus": df["fecha_dus"],
                    "agencia": df["aga"],
                })

                output_seca_cons = BytesIO()
                df_consolidado.to_excel(output_seca_cons, index=False, engine='openpyxl')
                output_seca_cons.seek(0)
                archivos_output.append(("CMPC_Madera_Seca_Consolidado.xlsx", output_seca_cons))

        # SUB-PROCESO 2: MADERA VERDE
        remate_verde = remate[remate["producto"] == "MADERA VERDE"].copy()
        
        if not remate_verde.empty:
            try:
                remate_verde['Desc_Carga_Calc'] = remate_verde['cant_piezas'].astype(str) + " PIECES, CHILEAN RADIATA PINE"
                
                contenedores_unicos_v = remate_verde['CONTENEDORREM'].unique()
                mapa_nota_v = {cnt: i+1 for i, cnt in enumerate(contenedores_unicos_v)}

                df_remate_extra_verde = pd.DataFrame({
                    "Nota": remate_verde['CONTENEDORREM'].map(mapa_nota_v),
                    "Venta": remate_verde["pedido"],
                    "Reserva": remate_verde["reserva"],
                    "Contenedor": remate_verde["CONTENEDORREM"],
                    "Sello Naviera (Carrier Seal)": remate_verde["sello_linea"],
                    "Descripción de la Carga": remate_verde["Desc_Carga_Calc"],
                    "N° de Pqts.": remate_verde["cant_paquetes"],
                    "Tara del Contenedor": remate_verde["tara"],
                    "Volumen Bruto de la Carga": remate_verde["volumen"],
                    "Peso Bruto de la Carga (documental)": remate_verde["neto"],
                    "Volumen Bruto del Contenedor": remate_verde["volumen"],
                    "Comentarios del Contenedor": remate_verde["pto_final"]
                })
                
                output_verde_remate = BytesIO()
                df_remate_extra_verde.to_excel(output_verde_remate, index=False, engine='openpyxl')
                output_verde_remate.seek(0)
                archivos_output.append(("Remate_CMPC_Madera_Verde.xlsx", output_verde_remate))
                
            except Exception as e:
                st.warning(f"Error generando Remate Extra Verde: {e}")

            # CONSOLIDADO VERDE
            tools_filtrado_v = tools[tools['CONTENEDORINF'].isin(remate_verde['CONTENEDORREM'])]
            remate_matched_v = remate_verde.set_index("CONTENEDORREM")
            tools_matched_v = tools_filtrado_v.set_index("CONTENEDORINF")
            
            df_v = tools_matched_v.join(remate_matched_v, how="left", rsuffix="_rem")
            
            if not df_v.empty:
                df_v[['contrato', 'item']] = df_v['Orden_Pedido'].astype(str).str.split('-', n=1, expand=True)
                df_v['fecha_dus'] = pd.to_datetime(df_v['fecha_aceptacion'], errors='coerce').dt.strftime('%d/%m/%Y')

                df_consolidado_v = pd.DataFrame({
                    "Npaquete": df_v["Nro_Paquete"],
                    "Contenedor": df_v.index,
                    "Sello": df_v["sello_linea"],
                    "Tara": df_v["tara"],
                    "Dimension": df_v["medida"],
                    "Tipo Cont.": df_v["tipo"],
                    "Directo": "N",
                    "Destino": df_v["pto_final"],
                    "Paquetes": "1",
                    "contrato": df_v["contrato"],
                    "item": df_v["item"],
                    "Naviera": df_v["linea"],
                    "Bodega": "",
                    "ubicación": "",
                    "Reserva": df_v["reserva"],
                    "Dus": df_v["dus"],
                    "fecha dus": df_v["fecha_dus"],
                    "agencia": df_v["aga"],
                })

                output_verde_cons = BytesIO()
                df_consolidado_v.to_excel(output_verde_cons, index=False, engine='openpyxl')
                output_verde_cons.seek(0)
                archivos_output.append(("CMPC_Madera_Verde_Consolidado.xlsx", output_verde_cons))

        if not archivos_output:
            return True, "Proceso finalizado, pero no se generaron archivos.", []

        return True, "Archivos generados exitosamente", archivos_output

    except Exception as e:
        st.error(f"Error en procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, str(e), []

# ==========================================
#      LÓGICA CMPC PAPEL (FINAL - NOTA POR CONTENEDOR)
# ==========================================
def procesar_cmpc_papel(rutas):
    st.info("Iniciando procesamiento CMPC Papel...")
    try:
        remate = pd.read_excel(rutas['remate'])
        tools = pd.read_excel(rutas['tools'])

        archivos_output = []

        # 1. NORMALIZACIÓN DE COLUMNAS Y CONTENEDORES
        remate.columns = [c.strip() for c in remate.columns]
        
        col_tara_rem = next((c for c in remate.columns if c.lower() == 'tara'), 'tara')
        col_pto_rem = next((c for c in remate.columns if c.lower() in ['pto_descarga', 'pto_final', 'puerto_destino']), 'pto_descarga')
        
        remate['sigla_cnt'] = remate['sigla_cnt'].astype(str).str.strip()
        remate['nro_cnt'] = remate['nro_cnt'].astype(str).str.strip()
        remate['dv_cnt'] = remate['dv_cnt'].astype(str).str.strip()

        def construir_contenedor_rem(row):
            sigla = str(row['sigla_cnt']).strip()
            val_num = str(row['nro_cnt']).split('.')[0].strip()
            dv = str(row['dv_cnt']).strip()
            return f"{sigla}-{val_num.zfill(6)}-{dv}"

        remate['CONTENEDORREM'] = remate.apply(construir_contenedor_rem, axis=1)

        tools.columns = [c.strip() for c in tools.columns]
        
        tools['Cnt_Sigla'] = tools['Cnt_Sigla'].astype(str).str.strip()
        tools['Cnt_Nro'] = tools['Cnt_Nro'].astype(str).str.strip()
        tools['Cnt_DV'] = tools['Cnt_DV'].astype(str).str.strip()
        
        col_sello_tools = next((c for c in tools.columns if c.lower() == 'sello_linea'), 'Sello_linea')
        if col_sello_tools in tools.columns:
            tools["Sello_linea_clean"] = tools[col_sello_tools].astype(str).str.replace("-", "", regex=False).str.strip()
        else:
            tools["Sello_linea_clean"] = ""

        col_peso_tools = next((c for c in tools.columns if c.lower() == 'peso_lote'), None)
        if col_peso_tools:
            tools[col_peso_tools] = tools[col_peso_tools].astype(str).str.replace(',', '.', regex=False)
            tools[col_peso_tools] = pd.to_numeric(tools[col_peso_tools], errors='coerce').fillna(0)
        else:
            tools['Peso_lote'] = 0
            col_peso_tools = 'Peso_lote'

        def construir_contenedor_tools(row):
            sigla = str(row['Cnt_Sigla']).strip()
            val_num = str(row['Cnt_Nro']).split('.')[0].strip()
            dv = str(row['Cnt_DV']).strip()
            return f"{sigla}-{val_num.zfill(6)}-{dv}"

        tools['CONTENEDORINF'] = tools.apply(construir_contenedor_tools, axis=1)

        # 2. GENERAR ARCHIVO NUEVO "REMATE_CMPC_PAPEL"
        try:
            grupo_tools = tools.groupby(['Orden_Pedido', 'CONTENEDORINF']).agg({
                'Reserva': 'first',
                'Sello_linea_clean': 'first',
                'Nro_Paquete': 'count',
                col_peso_tools: 'sum'
            }).reset_index()

            remate_subset = remate[['CONTENEDORREM', col_tara_rem, col_pto_rem]].drop_duplicates('CONTENEDORREM')
            
            df_nuevo = grupo_tools.merge(
                remate_subset,
                left_on='CONTENEDORINF',
                right_on='CONTENEDORREM',
                how='left'
            )

            contenedores_unicos = df_nuevo['CONTENEDORINF'].unique()
            mapa_id_contenedor = {cnt: i+1 for i, cnt in enumerate(contenedores_unicos)}
            
            df_exportar = pd.DataFrame()
            df_exportar['Nota'] = df_nuevo['CONTENEDORINF'].map(mapa_id_contenedor)
            df_exportar['Número Venta'] = df_nuevo['Orden_Pedido']
            df_exportar['Reserva'] = df_nuevo['Reserva']
            df_exportar['Contenedor'] = df_nuevo['CONTENEDORINF']
            df_exportar['Sello Naviera (Carrier Seal)'] = df_nuevo['Sello_linea_clean']
            df_exportar['Descripción de la Carga'] = "PAPEL KRAFT"
            df_exportar['N° de Pqts.'] = df_nuevo['Nro_Paquete']
            df_exportar['Tara del Contenedor'] = df_nuevo[col_tara_rem]
            df_exportar['Peso Bruto de la Carga (documental)'] = df_nuevo[col_peso_tools]
            df_exportar['Comentarios del Contenedor'] = df_nuevo[col_pto_rem]

            output_remate = BytesIO()
            df_exportar.to_excel(output_remate, index=False, engine='openpyxl')
            output_remate.seek(0)
            archivos_output.append(("Remate_CMPC_Papel.xlsx", output_remate))

        except Exception as e:
            st.warning(f"Error generando Remate Nuevo: {e}")

        # 3. GENERAR ARCHIVO ANTIGUO "CONSOLIDADO"
        try:
            remate_papel = remate[remate["producto"] == "PAPEL KRAFT"].copy()
            tools_filt = tools[tools['CONTENEDORINF'].isin(remate_papel['CONTENEDORREM'])].copy()
            
            df_cons = tools_filt.set_index("CONTENEDORINF").join(
                remate_papel.set_index("CONTENEDORREM"), 
                how="left", 
                rsuffix="_rem"
            )

            if not df_cons.empty:
                df_cons['fecha_dus'] = pd.to_datetime(df_cons['fecha_aceptacion'], errors='coerce').dt.strftime('%d/%m/%Y')
                
                df_consolidado_final = pd.DataFrame({
                    "Etiqueta": df_cons["Nro_Paquete"], 
                    "Contenedor": df_cons.index, 
                    "Sello": df_cons["sello_linea"],
                    "Tara": df_cons[col_tara_rem], 
                    "Dimension": df_cons["medida"], 
                    "Tipo Cont.": df_cons["tipo"], 
                    "Directo": "N",
                    "Destino": df_cons["pto_final"], 
                    "Fardos": "1", 
                    "contrato": df_cons["Orden_Pedido"], 
                    "item": "10",
                    "Naviera": df_cons["linea"], 
                    "Bodega": "", 
                    "ubicación": "", 
                    "Reserva": df_cons["reserva"], 
                    "Dus": df_cons["dus"], 
                    "fecha dus": df_cons["fecha_dus"], 
                    "agencia": df_cons["aga"]
                })

                output_consolidado = BytesIO()
                df_consolidado_final.to_excel(output_consolidado, index=False, engine='openpyxl')
                output_consolidado.seek(0)
                archivos_output.append(("CMPC_Papel_Consolidado.xlsx", output_consolidado))

        except Exception as e:
            st.warning(f"Error generando Consolidado: {e}")

        if not archivos_output:
            return True, "Proceso finalizado, pero no se generaron archivos.", []

        return True, "Archivos generados exitosamente", archivos_output

    except Exception as e:
        st.error(f"Error en procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, str(e), []

# ==========================================
#      LÓGICA CMPC PLYWOOD (FINAL - NOTA POR CONTENEDOR)
# ==========================================
def procesar_cmpc_plywood(rutas):
    st.info("Iniciando procesamiento CMPC Plywood...")
    try:
        remate = pd.read_excel(rutas['remate'])
        tools = pd.read_excel(rutas['tools'])

        remate['sigla_cnt'] = remate['sigla_cnt'].astype(str).str.strip()
        remate['nro_cnt'] = remate['nro_cnt'].astype(str).str.strip()
        remate['dv_cnt'] = remate['dv_cnt'].astype(str).str.strip()
        
        tools['Cnt_Sigla'] = tools['Cnt_Sigla'].astype(str).str.strip()
        tools['Cnt_Nro'] = tools['Cnt_Nro'].astype(str).str.strip()
        tools['Cnt_DV'] = tools['Cnt_DV'].astype(str).str.strip()
        
        if "Sello_linea" in tools.columns:
            tools["Sello_linea_clean"] = tools["Sello_linea"].astype(str).str.replace("-", "", regex=False).str.strip()

        def construir_contenedor(row):
            sigla = str(row['sigla_cnt']).strip()
            val_num = str(row['nro_cnt'])
            if '.' in val_num: numero = val_num.split('.')[0].strip()
            else: numero = val_num.strip()
            dv = str(row['dv_cnt']).strip()
            numero = numero.zfill(6)
            contenedor = f"{sigla}-{numero}-{dv}"
            return contenedor

        def construir_contenedor2(row):
            sigla = str(row['Cnt_Sigla']).strip()
            val_num = str(row['Cnt_Nro'])
            if '.' in val_num: numero = val_num.split('.')[0].strip()
            else: numero = val_num.strip()
            dv = str(row['Cnt_DV']).strip()
            numero = numero.zfill(6)
            contenedor = f"{sigla}-{numero}-{dv}"
            return contenedor

        remate['CONTENEDORREM'] = remate.apply(construir_contenedor, axis=1)
        tools['CONTENEDORINF'] = tools.apply(construir_contenedor2, axis=1)

        archivos_output = []

        remate_ply = remate[remate["producto"] == "PLYWOOD"].copy()
        
        if remate_ply.empty: 
            return True, "No se encontraron registros con producto 'PLYWOOD' en el archivo Remate.", []

        # GENERAR REMATE EXTRA
        try:
            remate_ply['Desc_Carga_Calc'] = remate_ply['cant_piezas'].astype(str) + " PIECES, PLYWOOD"
            
            contenedores_unicos = remate_ply['CONTENEDORREM'].unique()
            mapa_nota = {cnt: i+1 for i, cnt in enumerate(contenedores_unicos)}
            
            df_remate_extra = pd.DataFrame({
                "Nota": remate_ply['CONTENEDORREM'].map(mapa_nota),
                "Venta": remate_ply["pedido"],
                "Reserva": remate_ply["reserva"],
                "Contenedor": remate_ply["CONTENEDORREM"],
                "Sello Naviera (Carrier Seal)": remate_ply["sello_linea"],
                "Descripción de la Carga": remate_ply['Desc_Carga_Calc'],
                "N° de Pqts.": remate_ply["cant_paquetes"],
                "Tara del Contenedor": remate_ply["tara"],
                "Volumen Bruto de la Carga": remate_ply["volumen"],
                "Peso Bruto de la Carga (documental)": remate_ply["neto"],
                "Volumen Bruto del Contenedor": remate_ply["volumen"],
                "Comentarios del Contenedor": remate_ply["pto_final"]
            })
            
            output_remate = BytesIO()
            df_remate_extra.to_excel(output_remate, index=False, engine='openpyxl')
            output_remate.seek(0)
            archivos_output.append(("Remate_CMPC_Plywood.xlsx", output_remate))
            
        except Exception as e:
            st.warning(f"Error generando Remate Extra Plywood: {e}")

        # LÓGICA ORIGINAL: CONSOLIDADO
        tools_filt = tools[tools['CONTENEDORINF'].isin(remate_ply['CONTENEDORREM'])]
        
        df = tools_filt.set_index("CONTENEDORINF").join(remate_ply.set_index("CONTENEDORREM"), how="left", rsuffix="_rem")
        
        if not df.empty:
            df['fecha_dus'] = pd.to_datetime(df['fecha_aceptacion'], errors='coerce').dt.strftime('%d/%m/%Y')
            
            df_consolidado = pd.DataFrame({
                "Npaquete": df["Nro_Paquete"], 
                "Contenedor": df.index, 
                "Sello": df["sello_linea"],
                "Tara": df["tara"], 
                "Dimension": df["medida"], 
                "Tipo Cont.": df["tipo"], 
                "Directo": "N",
                "Destino": df["pto_final"], 
                "Fardos": "1", 
                "contrato": df["Orden_Pedido"], 
                "item": "10",
                "Naviera": df["linea"], 
                "Bodega": "", 
                "ubicación": "", 
                "Reserva": df["reserva"], 
                "Dus": df["dus"], 
                "fecha dus": df["fecha_dus"], 
                "agencia": df["aga"]
            })

            output_consolidado = BytesIO()
            df_consolidado.to_excel(output_consolidado, index=False, engine='openpyxl')
            output_consolidado.seek(0)
            archivos_output.append(("CMPC_Plywood_Consolidado.xlsx", output_consolidado))

        if not archivos_output:
            return True, "Proceso finalizado sin generar archivos.", []

        return True, "Archivos generados exitosamente", archivos_output

    except Exception as e:
        st.error(f"Error en procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, str(e), []

# ==========================================
#      INTERFAZ STREAMLIT
# ==========================================
CONFIG_ARCHIVOS = {
    "Madera": [
        {"id": "programa", "nombre": "Programa", "opcional": False},
        {"id": "saldos",   "nombre": "Saldos",   "opcional": True},
        {"id": "historico","nombre": "Remates Ant.", "opcional": True, "multiple": True},
        {"id": "despacho", "nombre": "Despacho", "opcional": False},
        {"id": "detalle",  "nombre": "Detalle",  "opcional": False},
        {"id": "informe",  "nombre": "Informe",  "opcional": False},
        {"id": "zoopp",    "nombre": "Zoopp",    "opcional": False},
    ],
    "Celulosa BKP EKP UKP": [
        {"id": "programa", "nombre": "Programa", "opcional": False},
        {"id": "saldos",   "nombre": "Saldos",   "opcional": True},
        {"id": "tools",    "nombre": "Tools",    "opcional": False},
        {"id": "historico","nombre": "Remates Ant.", "opcional": True, "multiple": True},
    ],
    "Celulosa DP": [
        {"id": "programa", "nombre": "Programa", "opcional": False},
        {"id": "saldos",   "nombre": "Saldos",   "opcional": True},
        {"id": "informe",  "nombre": "Informe",  "opcional": False},
        {"id": "historico","nombre": "Remates Ant.", "opcional": True, "multiple": True},
    ],
    "SAG": [
        {"id": "remate", "nombre": "Remate", "opcional": False},
        {"id": "picking", "nombre": "Picking", "opcional": False},
        {"id": "sag",   "nombre": "SIF",   "opcional": False, "multiple": True}
    ],
    "CMPC Celulosa": [
        {"id": "remate", "nombre": "Remate", "opcional": False},
        {"id": "tools",  "nombre": "Tools",  "opcional": False},
    ],
    "CMPC Madera": [
        {"id": "remate", "nombre": "Remate", "opcional": False},
        {"id": "informe", "nombre": "Tools", "opcional": False},
    ],
    "CMPC Papel": [
        {"id": "remate", "nombre": "Remate", "opcional": False},
        {"id": "tools",  "nombre": "Tools",  "opcional": False},
    ],
    "CMPC Plywood": [
        {"id": "remate", "nombre": "Remate", "opcional": False},
        {"id": "tools",  "nombre": "Tools",  "opcional": False},
    ]        
}

def get_file_uploader_key(file_id, session_id):
    return f"{file_id}_{session_id}"

def aplicar_estilos():
    st.markdown("""
        <style>
        /* Animaciones y estilo para los botones normales */
        div.stButton > button:first-child {
            border-radius: 10px;
            font-weight: 600;
            transition: all 0.3s ease-in-out;
            border: 1px solid #d1d5db;
        }
        div.stButton > button:first-child:hover {
            border-color: #3b82f6;
            color: #3b82f6;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            transform: translateY(-2px);
        }
        
        /* Estilo especial para el botón de "🚀 Ejecutar Proceso" (Form Submit) */
        div.stFormSubmitButton > button:first-child {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
            color: white;
            border: none;
            border-radius: 10px;
            font-weight: bold;
            transition: all 0.3s ease;
        }
        div.stFormSubmitButton > button:first-child:hover {
            background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
            box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.4);
            transform: translateY(-2px);
            color: white;
        }

        /* Estilo premium para los botones de descarga */
        div.stDownloadButton > button:first-child {
            background-color: #10b981; /* Verde esmeralda */
            color: white;
            border-radius: 8px;
            font-weight: 600;
            border: none;
            width: 100%;
            transition: all 0.3s ease;
        }
        div.stDownloadButton > button:first-child:hover {
            background-color: #059669;
            box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.4);
            transform: translateY(-3px);
            color: white;
        }
        
        /* Ajuste de las tarjetas de subida de archivos */
        section[data-testid="stFileUploadDropzone"] {
            border-radius: 12px;
            border: 2px dashed #cbd5e1;
            background-color: #f8fafc;
            transition: all 0.3s;
        }
        section[data-testid="stFileUploadDropzone"]:hover {
            border-color: #3b82f6;
            background-color: #eff6ff;
        }
        </style>
    """, unsafe_allow_html=True)

def main():
    st.set_page_config(
        page_title="Agente CFS",
        page_icon="📦",
        layout="wide"
    )
    
    # Llamamos a los estilos mágicos aquí
    aplicar_estilos()
    
    # Inicializar session state
    if 'empresa_seleccionada' not in st.session_state:
        st.session_state.empresa_seleccionada = None
    if 'tipo_material' not in st.session_state:
        st.session_state.tipo_material = None
    if 'archivos_cargados' not in st.session_state:
        st.session_state.archivos_cargados = {}
    if 'session_id' not in st.session_state:
        st.session_state.session_id = str(hash(str(datetime.datetime.now())))
    
    st.title("📊 Agente CFS")
    st.markdown("---") # Una línea divisoria elegante
    
    # Mostrar pantalla inicial si no hay empresa seleccionada
    if st.session_state.empresa_seleccionada is None:
        mostrar_inicio_empresas()
    else:
        if st.session_state.tipo_material is None:
            if st.session_state.empresa_seleccionada == "Arauco":
                mostrar_menu_materiales_arauco()
            else:
                mostrar_menu_materiales_cmpc()
        else:
            mostrar_panel_proceso()

def mostrar_inicio_empresas():
    st.header("Seleccione Empresa")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("**ARAUCO**", use_container_width=True, type="primary"):
            st.session_state.empresa_seleccionada = "Arauco"
            st.session_state.tipo_material = None
            st.rerun()
    
    with col2:
        if st.button("**CMPC**", use_container_width=True, type="primary"):
            st.session_state.empresa_seleccionada = "CMPC"
            st.session_state.tipo_material = None
            st.rerun()

def mostrar_menu_materiales_arauco():
    st.header("Arauco - Seleccione Material")
    
    if st.button("← Volver a Empresas"):
        st.session_state.empresa_seleccionada = None
        st.session_state.tipo_material = None
        st.rerun()
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("**Celulosa DP**", use_container_width=True):
            st.session_state.tipo_material = "Celulosa DP"
            st.rerun()
        
        if st.button("**Madera**", use_container_width=True):
            st.session_state.tipo_material = "Madera"
            st.rerun()
    
    with col2:
        if st.button("**Celulosa BKP EKP UKP**", use_container_width=True):
            st.session_state.tipo_material = "Celulosa BKP EKP UKP"
            st.rerun()
        
        if st.button("**SAG**", use_container_width=True):
            st.session_state.tipo_material = "SAG"
            st.rerun()

def mostrar_menu_materiales_cmpc():
    st.header("CMPC - Seleccione Material")
    
    if st.button("← Volver a Empresas"):
        st.session_state.empresa_seleccionada = None
        st.session_state.tipo_material = None
        st.rerun()
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("**Celulosa**", use_container_width=True):
            st.session_state.tipo_material = "CMPC Celulosa"
            st.rerun()
        
        if st.button("**Papel**", use_container_width=True):
            st.session_state.tipo_material = "CMPC Papel"
            st.rerun()
    
    with col2:
        if st.button("**Madera**", use_container_width=True):
            st.session_state.tipo_material = "CMPC Madera"
            st.rerun()
        
        if st.button("**Plywood**", use_container_width=True):
            st.session_state.tipo_material = "CMPC Plywood"
            st.rerun()

def mostrar_panel_proceso():
    st.header(f"Panel: {st.session_state.tipo_material}")
    
    if st.button("← Volver"):
        st.session_state.tipo_material = None
        st.session_state.archivos_generados = None # Limpiar al volver
        st.rerun()
    
    lista_archivos = CONFIG_ARCHIVOS.get(st.session_state.tipo_material, [])
    
    # Limpiar archivos cargados si cambió el tipo de material
    if 'last_material' not in st.session_state or st.session_state.last_material != st.session_state.tipo_material:
        st.session_state.archivos_cargados = {}
        st.session_state.archivos_generados = None # Asegurarnos de limpiar salidas previas
        st.session_state.last_material = st.session_state.tipo_material
    
    st.subheader("Carga de Archivos")
    
    # Crear formulario para subir archivos
    with st.form("upload_form"):
        for item in lista_archivos:
            es_multiple = item.get("multiple", False)
            required = "" if item["opcional"] else "🔴"
            
            if es_multiple:
                uploaded_files = st.file_uploader(
                    f"{required} {item['nombre']} {'(Múltiple)' if es_multiple else ''}",
                    type=['xlsx', 'xls', 'dbf'],
                    accept_multiple_files=True,
                    key=get_file_uploader_key(item["id"], st.session_state.session_id)
                )
                if uploaded_files:
                    temp_files = []
                    for uploaded_file in uploaded_files:
                        # Guardar archivo temporal
                        with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{uploaded_file.name}") as tmp_file:
                            tmp_file.write(uploaded_file.getvalue())
                            temp_files.append(tmp_file.name)
                    st.session_state.archivos_cargados[item["id"]] = temp_files
                    st.success(f"{len(uploaded_files)} archivo(s) cargado(s)")
            else:
                uploaded_file = st.file_uploader(
                    f"{required} {item['nombre']}",
                    type=['xlsx', 'xls', 'dbf'],
                    key=get_file_uploader_key(item["id"], st.session_state.session_id)
                )
                if uploaded_file:
                    # Guardar archivo temporal
                    with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{uploaded_file.name}") as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        st.session_state.archivos_cargados[item["id"]] = tmp_file.name
                    st.success(f"Archivo cargado: {uploaded_file.name}")
        
        submit_button = st.form_submit_button("🚀 Ejecutar Proceso")
    
    if submit_button:
        ejecutar_proceso()

    # --- AQUÍ ESTÁ LA MAGIA ---
    # Mostramos los botones FUERA del formulario y basados en session_state
    if st.session_state.get('archivos_generados'):
        st.subheader("📥 Archivos Generados")
        cols = st.columns(min(3, len(st.session_state.archivos_generados)))
        
        for idx, (nombre, archivo_bytes) in enumerate(st.session_state.archivos_generados):
            with cols[idx % 3]:
                st.download_button(
                    label=f"Descargar {nombre}",
                    data=archivo_bytes,
                    file_name=nombre,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"btn_descarga_{nombre}" # Es clave darle un ID único a cada botón
                )

import time # Asegúrate de tener 'import time' al inicio de tu app.py si no lo tienes

def ejecutar_proceso():
    # Validar archivos obligatorios
    lista_archivos = CONFIG_ARCHIVOS.get(st.session_state.tipo_material, [])
    faltantes = []
    
    for item in lista_archivos:
        if not item["opcional"] and item["id"] not in st.session_state.archivos_cargados:
            faltantes.append(item["nombre"])
    
    if faltantes:
        st.error(f"Faltan archivos obligatorios:\n- " + "\n- ".join(faltantes))
        return
    
    # Seleccionar lógica según tipo de material
    tipo_material = st.session_state.tipo_material
    rutas = st.session_state.archivos_cargados
    
    # ==========================================
    # PANTALLA DE CARGA ESTÉTICA
    # ==========================================
    with st.status("🚀 Iniciando procesamiento...", expanded=True) as status:
        st.write("📂 Leyendo archivos Excel...")
        time.sleep(0.3) # Pequeña pausa para que la animación se aprecie
        
        st.write("⚙️ Cruzando información y aplicando lógica de negocio...")
        
        # Aquí corre tu código pesado
        if tipo_material == "Madera":
            exito, mensaje, archivos = procesar_madera(rutas)
        elif tipo_material == "Celulosa BKP EKP UKP":
            exito, mensaje, archivos = procesar_celulosa_cb(rutas)
        elif tipo_material == "Celulosa DP":
            exito, mensaje, archivos = procesar_celulosa_sb(rutas)
        elif tipo_material == "SAG":
            exito, mensaje, archivos = procesar_sag(rutas)
        elif tipo_material == "CMPC Celulosa":
            exito, mensaje, archivos = procesar_cmpc_celulosa(rutas)
        elif tipo_material == "CMPC Madera":
            exito, mensaje, archivos = procesar_cmpc_madera(rutas)
        elif tipo_material == "CMPC Papel":
            exito, mensaje, archivos = procesar_cmpc_papel(rutas)
        elif tipo_material == "CMPC Plywood":
            exito, mensaje, archivos = procesar_cmpc_plywood(rutas)
        else:
            exito, mensaje, archivos = False, "Lógica no implementada", []
        
        st.write("📝 Generando reportes de salida...")
        
        # Actualizamos el estado de la cajita dependiendo del resultado
        if exito:
            status.update(label="¡Procesamiento Completado!", state="complete", expanded=False)
        else:
            status.update(label="Ocurrió un error en el proceso", state="error", expanded=True)

    # Mostrar resultados y animaciones
    if exito:
        st.toast('¡Archivos generados con éxito!', icon='🎉') # Notificación flotante
        # st.balloons() # Descomenta esto si quieres globos volando por la pantalla (a veces es mucho, pero es divertido)
        
        st.success(mensaje)
        st.session_state.archivos_generados = archivos
    else:
        st.session_state.archivos_generados = None
        st.error(f"Error: {mensaje}")

if __name__ == "__main__":

    main()
