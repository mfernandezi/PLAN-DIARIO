"""
Dashboard CSV - Exporta KPI de Perforación a CSV legible.

Genera archivos CSV planos a partir del Excel KPI PyT,
listos para usar en dashboards, Power BI, o análisis.

Uso:
    python dashboard_csv.py                  # Exporta todo
    python dashboard_csv.py 15-01 16-01      # Solo esas fechas
    python dashboard_csv.py --output mi_kpi  # Nombre personalizado
"""

import csv
import sys
import os
from leer_excel import (
    cargar_excel,
    obtener_dia,
    EXCEL_PATH,
    HOJAS_EXCLUIDAS,
    DiaData,
)
import openpyxl


def dia_a_filas(dia: DiaData) -> list[dict]:
    """Convierte un DiaData en filas planas (1 fila por equipo)."""
    filas = []

    for fase_label in ["F12", "F10", "F09", "F11"]:
        fase = dia.fases.get(fase_label)
        if not fase:
            continue

        for eq in fase.equipos:
            filas.append({
                "Fecha": dia.fecha,
                "Fase": fase_label,
                "Equipo": eq.equipo,
                "Turno_A": eq.turno_a,
                "Turno_B": eq.turno_b,
                "Total": eq.total,
                "Plan": eq.plan,
                "Cumplimiento_TA": round(eq.cumplimiento_ta * 100, 1) if eq.cumplimiento_ta else 0.0,
                "Cumplimiento_TB": round(eq.cumplimiento_tb * 100, 1) if eq.cumplimiento_tb else 0.0,
                "Cumplimiento_Diario": round(eq.cumplimiento_diario * 100, 1) if eq.cumplimiento_diario else 0.0,
                "Estado_TA": eq.estado_ta,
                "Estado_TB": eq.estado_tb,
                "Tipo": "Equipo",
            })

        # Fila total de la fase
        filas.append({
            "Fecha": dia.fecha,
            "Fase": fase_label,
            "Equipo": "TOTAL_FASE",
            "Turno_A": fase.total_turno_a,
            "Turno_B": fase.total_turno_b,
            "Total": fase.total,
            "Plan": fase.plan,
            "Cumplimiento_TA": round(fase.cumplimiento_ta * 100, 1) if fase.cumplimiento_ta else 0.0,
            "Cumplimiento_TB": round(fase.cumplimiento_tb * 100, 1) if fase.cumplimiento_tb else 0.0,
            "Cumplimiento_Diario": round(fase.cumplimiento_diario * 100, 1) if fase.cumplimiento_diario else 0.0,
            "Estado_TA": "",
            "Estado_TB": "",
            "Tipo": "Total_Fase",
        })

    # Fila TOTAL METROS del día
    filas.append({
        "Fecha": dia.fecha,
        "Fase": "TODAS",
        "Equipo": "TOTAL_DIA",
        "Turno_A": dia.total_turno_a,
        "Turno_B": dia.total_turno_b,
        "Total": dia.total_metros,
        "Plan": dia.plan_total,
        "Cumplimiento_TA": round(dia.cumplimiento_ta * 100, 1) if dia.cumplimiento_ta else 0.0,
        "Cumplimiento_TB": round(dia.cumplimiento_tb * 100, 1) if dia.cumplimiento_tb else 0.0,
        "Cumplimiento_Diario": round(dia.cumplimiento_diario * 100, 1) if dia.cumplimiento_diario else 0.0,
        "Estado_TA": "",
        "Estado_TB": "",
        "Tipo": "Total_Dia",
    })

    return filas


def dia_a_filas_roc(dia: DiaData) -> list[dict]:
    """Genera filas de Metros ROC para un día."""
    roc = dia.metros_roc
    filas = []

    for fase_nombre, valor in [
        ("F09", roc.fase_9),
        ("F10", roc.fase_10),
        ("F11", roc.fase_11),
        ("F12", roc.fase_12),
    ]:
        filas.append({
            "Fecha": dia.fecha,
            "Fase": fase_nombre,
            "Metros_ROC": valor,
        })

    filas.append({
        "Fecha": dia.fecha,
        "Fase": "TOTAL_ROC",
        "Metros_ROC": roc.total,
        "Turno_A_ROC": roc.turno_a,
        "Turno_B_ROC": roc.turno_b,
        "Plan_Diario_ROC": roc.plan_diario,
    })

    return filas


def exportar_metros_csv(dias: dict, output: str = "kpi_metros"):
    """Exporta metros de perforación a CSV."""
    archivo = f"{output}.csv"
    todas_filas = []

    for fecha in sorted(dias.keys()):
        todas_filas.extend(dia_a_filas(dias[fecha]))

    if not todas_filas:
        print("No hay datos para exportar.")
        return None

    columnas = [
        "Fecha", "Fase", "Equipo", "Turno_A", "Turno_B", "Total",
        "Plan", "Cumplimiento_TA", "Cumplimiento_TB", "Cumplimiento_Diario",
        "Estado_TA", "Estado_TB", "Tipo",
    ]

    with open(archivo, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=columnas, delimiter=";")
        writer.writeheader()
        writer.writerows(todas_filas)

    print(f"  Metros exportados: {archivo} ({len(todas_filas)} filas)")
    return archivo


def exportar_roc_csv(dias: dict, output: str = "kpi_roc"):
    """Exporta metros ROC a CSV."""
    archivo = f"{output}.csv"
    todas_filas = []

    for fecha in sorted(dias.keys()):
        todas_filas.extend(dia_a_filas_roc(dias[fecha]))

    if not todas_filas:
        print("No hay datos ROC para exportar.")
        return None

    columnas = [
        "Fecha", "Fase", "Metros_ROC", "Turno_A_ROC", "Turno_B_ROC", "Plan_Diario_ROC",
    ]

    with open(archivo, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=columnas, delimiter=";", extrasaction="ignore")
        writer.writeheader()
        writer.writerows(todas_filas)

    print(f"  ROC exportados:   {archivo} ({len(todas_filas)} filas)")
    return archivo


def exportar_resumen_diario_csv(dias: dict, output: str = "kpi_resumen_diario"):
    """Exporta un resumen de 1 fila por día."""
    archivo = f"{output}.csv"
    filas = []

    for fecha in sorted(dias.keys()):
        dia = dias[fecha]
        fila = {
            "Fecha": dia.fecha,
            "Total_Turno_A": dia.total_turno_a,
            "Total_Turno_B": dia.total_turno_b,
            "Total_Metros": dia.total_metros,
            "Plan_Total": dia.plan_total,
            "Cumplimiento_Diario_%": round(dia.cumplimiento_diario * 100, 1) if dia.cumplimiento_diario else 0.0,
            "ROC_Total": dia.metros_roc.total,
            "ROC_Plan": dia.metros_roc.plan_diario,
        }
        # Agregar total por fase
        for fase_label in ["F09", "F10", "F11", "F12"]:
            fase = dia.fases.get(fase_label)
            if fase:
                fila[f"{fase_label}_Total"] = fase.total
                fila[f"{fase_label}_Plan"] = fase.plan
                fila[f"{fase_label}_Cump_%"] = round(fase.cumplimiento_diario * 100, 1) if fase.cumplimiento_diario else 0.0
            else:
                fila[f"{fase_label}_Total"] = 0.0
                fila[f"{fase_label}_Plan"] = 0.0
                fila[f"{fase_label}_Cump_%"] = 0.0

        filas.append(fila)

    if not filas:
        print("No hay datos para resumen.")
        return None

    columnas = list(filas[0].keys())

    with open(archivo, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=columnas, delimiter=";")
        writer.writeheader()
        writer.writerows(filas)

    print(f"  Resumen diario:   {archivo} ({len(filas)} filas)")
    return archivo


def main():
    args = sys.argv[1:]

    # Parsear argumento --output
    output_base = "kpi"
    if "--output" in args:
        idx = args.index("--output")
        if idx + 1 < len(args):
            output_base = args[idx + 1]
            args = args[:idx] + args[idx + 2:]
        else:
            args = args[:idx]

    fechas = [a for a in args if not a.startswith("--")]

    print(f"Leyendo Excel: {EXCEL_PATH}")

    if fechas:
        # Cargar solo las fechas solicitadas
        dias = {}
        for fecha in fechas:
            dia = obtener_dia(fecha)
            if dia:
                dias[fecha] = dia
            else:
                print(f"  AVISO: No se encontró la hoja '{fecha}'")
    else:
        # Cargar todo
        data = cargar_excel()
        dias = data["dias"]
        print(f"Hojas encontradas: {data['total_hojas']}")

    if not dias:
        print("No se encontraron datos. Verifica las fechas.")
        sys.exit(1)

    print(f"\nExportando {len(dias)} día(s)...\n")

    archivos = []
    a = exportar_metros_csv(dias, f"{output_base}_metros")
    if a:
        archivos.append(a)

    a = exportar_roc_csv(dias, f"{output_base}_roc")
    if a:
        archivos.append(a)

    a = exportar_resumen_diario_csv(dias, f"{output_base}_resumen_diario")
    if a:
        archivos.append(a)

    print(f"\nListo. {len(archivos)} archivo(s) CSV generados.")
    print("Separador: punto y coma (;) | Encoding: UTF-8 con BOM")
    print("Listos para importar en Power BI, Excel, o cualquier herramienta.")


if __name__ == "__main__":
    main()
