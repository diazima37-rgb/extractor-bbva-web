"""Módulo de extracción de movimientos BBVA.

Este módulo expone funciones para leer un estado de cuenta PDF y generar un
archivo Excel con los movimientos.

Usa Claude Vision (Anthropic API) para leer las imágenes del PDF y extraer
los movimientos en formato JSON. Esta es la solución más robusta para PDFs
escaneados y funciona en cualquier entorno (incluyendo Streamlit Cloud).

El código está pensado para ser usado desde una GUI (tkinter) u otro front-end.
"""

import base64
import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import certifi
import numpy as np
import pdfplumber
from pdf2image import convert_from_path
from anthropic import Anthropic

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# Tipos
Movimiento = Dict[str, Optional[str]]
InfoCuenta = Dict[str, Optional[str]]
LogFunc = Callable[[str], None]

# Patrones de reconocimiento
PATRON_FECHA = re.compile(r"(\d{1,2}/[A-Za-z]{3})")
PATRON_MONTO = re.compile(r"(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2}))")


def _norm_text(s: str) -> str:
    """Normaliza texto para comparaciones (sin acentos, mayúsculas)."""
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r'[^A-Za-z0-9 ]', '', s)
    return s.strip().upper()


def _normalizar_monto(texto: Optional[str]) -> Optional[float]:
    if texto is None:
        return None
    s = str(texto).strip()
    if not s:
        return None
    s = s.replace('$', '').replace(' ', '').replace('US$', '')
    # Convertir comas a punto si solo hay comas
    if s.count(',') == 1 and s.count('.') == 0:
        s = s.replace(',', '.')
    else:
        s = s.replace(',', '')
    try:
        return float(s)
    except Exception:
        return None


def _get_anthropic_client():
    """Obtiene el cliente de Anthropic con API key de variable de entorno."""
    api_key = os.getenv('ANTHROPIC_API_KEY')
    if not api_key:
        raise ValueError("❌ Variable de entorno ANTHROPIC_API_KEY no está configurada")
    return Anthropic(api_key=api_key)


def _extraer_movimientos_desde_imagen_claude(
    imagen_path: str, cliente: Anthropic, log: LogFunc
) -> Tuple[List[Movimiento], str]:
    """Envía una imagen de página PDF a Claude Vision y extrae los movimientos."""
    log("🔎 Enviando imagen a Claude Vision...")
    
    # Leer imagen y convertir a base64
    with open(imagen_path, "rb") as f:
        image_data = base64.standard_b64encode(f.read()).decode("utf-8")
    
    prompt = """Analiza esta imagen de un estado de cuenta bancario BBVA y extrae todos los movimientos.

Para cada movimiento, extrae:
- Fecha de operación (ej: 01/DIC)
- Fecha de liquidación (ej: 01/DIC)
- Descripción completa del movimiento
- Monto (número decimal)
- Tipo: "ABONO" o "CARGO"

Responde SOLO en este formato JSON, sin explicaciones adicionales:
{
  "movimientos": [
    {
      "fecha_oper": "01/DIC",
      "fecha_liq": "01/DIC",
      "descripcion": "DC MAYORISTA,SA DE C",
      "monto": 2454.18,
      "tipo": "ABONO"
    }
  ],
  "texto_completo": "todo el texto extraído de la página"
}

Si no hay movimientos, devuelve {"movimientos": [], "texto_completo": "..."}"""

    try:
        response = cliente.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=4096,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": image_data,
                            },
                        },
                        {
                            "type": "text",
                            "text": prompt
                        }
                    ],
                }
            ],
        )
        
        # Extraer JSON de la respuesta
        texto_respuesta = response.content[0].text
        
        # Intentar parsear JSON
        try:
            # Buscar JSON en la respuesta
            inicio = texto_respuesta.find('{')
            fin = texto_respuesta.rfind('}') + 1
            if inicio >= 0 and fin > inicio:
                json_str = texto_respuesta[inicio:fin]
                datos = json.loads(json_str)
                movimientos = []
                
                for mov in datos.get("movimientos", []):
                    monto = mov.get("monto")
                    cargo = None
                    abono = None
                    
                    tipo = mov.get("tipo", "ABONO").upper()
                    if tipo == "CARGO":
                        cargo = _normalizar_monto(monto)
                    else:
                        abono = _normalizar_monto(monto)
                    
                    movimientos.append({
                        "Fecha_Oper": mov.get("fecha_oper", ""),
                        "Fecha_Liq": mov.get("fecha_liq", ""),
                        "Descripcion": mov.get("descripcion", ""),
                        "Referencia": "",
                        "Cargo": cargo,
                        "Abono": abono,
                        "Saldo_Oper": None,
                        "Saldo_Liq": None,
                        "Tipo": tipo,
                    })
                
                texto_completo = datos.get("texto_completo", "")
                return movimientos, texto_completo
        except json.JSONDecodeError:
            log(f"⚠ Error parseando JSON de Claude: {texto_respuesta[:200]}")
            return [], ""
            
    except Exception as e:
        log(f"❌ Error llamando a Claude Vision: {e}")
        return [], ""
    
    return [], ""


def _parsear_encabezado(texto: str) -> InfoCuenta:
    """Extrae información del encabezado del estado de cuenta."""
    info: InfoCuenta = {
        'numero_cuenta': None,
        'periodo': None,
        'fecha_corte': None,
        'saldo_anterior': None,
        'saldo_final': None,
        'total_cargos': None,
        'total_abonos': None,
    }

    # Buscar número de cuenta
    match = re.search(r'CUENTA\s*[:\-]?\s*(\d{4,})', texto, re.IGNORECASE)
    if match:
        info['numero_cuenta'] = match.group(1).strip()

    # Período
    match = re.search(r'PERI[ÓO]DO\s*[:\-]?\s*([\d/\-\sA-Za-z]+)', texto, re.IGNORECASE)
    if match:
        info['periodo'] = match.group(1).strip()

    # Fecha de corte
    match = re.search(r'CORTE\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if match:
        info['fecha_corte'] = match.group(1).strip()

    # Saldos y totales
    match = re.search(r'SALDO\s+ANTERIOR\s*[:\-]?\s*\$?\s*([\d.,]+)', texto, re.IGNORECASE)
    if match:
        info['saldo_anterior'] = match.group(1).strip()

    match = re.search(r'SALDO\s+FINAL\s*[:\-]?\s*\$?\s*([\d.,]+)', texto, re.IGNORECASE)
    if match:
        info['saldo_final'] = match.group(1).strip()

    match = re.search(r'TOTAL\s+CARGOS\s*[:\-]?\s*\$?\s*([\d.,]+)', texto, re.IGNORECASE)
    if match:
        info['total_cargos'] = match.group(1).strip()

    match = re.search(r'TOTAL\s+ABONOS\s*[:\-]?\s*\$?\s*([\d.,]+)', texto, re.IGNORECASE)
    if match:
        info['total_abonos'] = match.group(1).strip()

    return info


def extraer_movimientos_desde_pdf(
    pdf_path: str, log: LogFunc = print
) -> Tuple[List[Movimiento], InfoCuenta]:
    """Extrae movimientos y encabezado desde un PDF usando Claude Vision."""
    movimientos: List[Movimiento] = []
    info_cuenta: InfoCuenta = {}

    ruta = Path(pdf_path)
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo PDF: {pdf_path}")

    log(f"📄 Abriendo PDF: {ruta.name}")
    
    try:
        cliente = _get_anthropic_client()
    except ValueError as e:
        log(f"❌ {e}")
        raise
    
    # Convertir PDF a imágenes
    try:
        log(f"📸 Convirtiendo PDF a imágenes...")
        # Intentar con rutas comunes de poppler
        poppler_paths = [
            '/usr/local/bin',
            '/opt/homebrew/bin',
            '/usr/bin',
            None  # Sistema PATH
        ]
        
        imagenes = None
        for poppler_path in poppler_paths:
            try:
                if poppler_path:
                    imagenes = convert_from_path(str(ruta), dpi=200, poppler_path=poppler_path)
                else:
                    imagenes = convert_from_path(str(ruta), dpi=200)
                break
            except Exception:
                continue
        
        if not imagenes:
            raise Exception("No se pudo convertir PDF: poppler no instalado")
        
        log(f"  Total de páginas: {len(imagenes)}")
    except Exception as e:
        log(f"❌ Error convirtiendo PDF: {e}")
        raise

    # Procesar cada página
    for idx, imagen in enumerate(imagenes, start=1):
        log(f"  Página {idx}/{len(imagenes)}")
        
        # Guardar imagen temporalmente
        temp_path = f"/tmp/bbva_page_{idx}.png"
        imagen.save(temp_path, "PNG")
        
        try:
            # Extraer movimientos de esta página con Claude
            movs, texto = _extraer_movimientos_desde_imagen_claude(temp_path, cliente, log)
            
            if idx == 1 and texto:
                info_cuenta = _parsear_encabezado(texto)
                log("    Encabezado parseado.")
            
            if movs:
                log(f"    ✓ {len(movs)} movimientos encontrados")
                movimientos.extend(movs)
            
        except Exception as e:
            log(f"    ⚠ Error procesando página {idx}: {e}")
        finally:
            # Limpiar imagen temporal
            try:
                Path(temp_path).unlink()
            except:
                pass

    log(f"✓ Movimientos encontrados: {len(movimientos)}")
    return movimientos, info_cuenta


def generar_excel_movimientos(
    movimientos: List[Movimiento],
    info_cuenta: InfoCuenta,
    ruta_salida: str,
    log: LogFunc = print,
) -> None:
    """Genera el archivo Excel con dos hojas: Movimientos y Resumen."""

    ruta = Path(ruta_salida)
    ruta.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Movimientos"

    encabezados = [
        'Fecha_Oper',
        'Fecha_Liq',
        'Descripción',
        'Referencia',
        'Cargo',
        'Abono',
        'Saldo_Oper',
        'Saldo_Liq',
        'Tipo',
    ]
    ws.append(encabezados)

    header_fill = PatternFill(start_color='003F87', end_color='003F87', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF')
    header_align = Alignment(horizontal='center', vertical='center')

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align

    for idx, mov in enumerate(movimientos, start=2):
        ws.append([
            mov.get('Fecha_Oper'),
            mov.get('Fecha_Liq'),
            mov.get('Descripcion'),
            mov.get('Referencia'),
            mov.get('Cargo'),
            mov.get('Abono'),
            mov.get('Saldo_Oper'),
            mov.get('Saldo_Liq'),
            mov.get('Tipo'),
        ])

        if idx % 2 == 0:
            fill = PatternFill(start_color='E8E8E8', end_color='E8E8E8', fill_type='solid')
            for cell in ws[idx]:
                cell.fill = fill

    total_cargos = sum(m.get('Cargo') or 0 for m in movimientos)
    total_abonos = sum(m.get('Abono') or 0 for m in movimientos)

    total_row = len(movimientos) + 3
    ws[f'A{total_row}'] = 'TOTALES'
    ws[f'E{total_row}'] = total_cargos
    ws[f'F{total_row}'] = total_abonos

    ws[f'E{total_row}'].number_format = '0.00'
    ws[f'F{total_row}'].number_format = '0.00'
    ws[f'A{total_row}'].font = Font(bold=True)

    # Autoajustar columnas
    anchos = [12, 12, 40, 18, 15, 15, 15, 15, 12]
    for i, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    # Hoja resumen
    ws2 = wb.create_sheet('Resumen')
    ws2['A1'] = 'RESUMEN DE CUENTA'
    ws2['A1'].font = Font(bold=True, size=14)

    ws2['A3'] = 'Número de Movimientos:'
    ws2['B3'] = len(movimientos)

    ws2['A4'] = 'Total Cargos:'
    ws2['B4'] = total_cargos
    ws2['B4'].number_format = '0.00'

    ws2['A5'] = 'Total Abonos:'
    ws2['B5'] = total_abonos
    ws2['B5'].number_format = '0.00'

    ws2['A7'] = 'Concepto'
    ws2['B7'] = 'Monto'
    ws2['A8'] = 'Cargos'
    ws2['B8'] = total_cargos
    ws2['A9'] = 'Abonos'
    ws2['B9'] = total_abonos

    chart = BarChart()
    chart.title = 'Cargos vs Abonos'

    # Usar Reference para la serie de datos
    data = Reference(ws2, min_col=2, min_row=7, max_row=9)
    chart.add_data(data, titles_from_data=True)
    ws2.add_chart(chart, 'D3')

    log(f"💾 Guardando Excel en: {ruta}")
    wb.save(str(ruta))


if __name__ == '__main__':
    # Prueba rápida en consola
    def _log(m):
        print(m)

    pdf = 'doc01266120250626153613.pdf'
    excel = 'movimientos_bbva.xlsx'
    movs, info = extraer_movimientos_desde_pdf(pdf, log=_log)
    generar_excel_movimientos(movs, info, excel, log=_log)
