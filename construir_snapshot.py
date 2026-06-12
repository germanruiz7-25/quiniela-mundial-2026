#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
  PASO PREVIO (correr UNA sola vez):  construir_snapshot.py
  --------------------------------------------------------------------------
  Lee tu Excel ORIGINAL (con las formulas ya calculadas por Excel) y guarda
  un 'snapshot.json' con:
     - el listado de partidos (fila, equipos, dia, hora)
     - los vaticinios de cada participante, ligados a la fila del partido

  Por que es necesario: openpyxl (Python) NO recalcula formulas de Excel.
  Una vez que el script guarda el .xlsm, los nombres y totales calculados
  por formula se pierden. Con el snapshot, el sistema ya no depende de las
  formulas: calcula los puntos por su cuenta y siempre funciona.

  Cuando volver a correrlo: solo si cambian los vaticinios de alguien o se
  agregan/editan participantes. Para actualizar marcadores NO hace falta.
"""
import openpyxl, datetime, json, re
from pathlib import Path

EXCEL_PATH = "2026_QUINIELA_CONSOLIDADO.xlsm"
HOJA_MARCADORES = "Marcadores_FINALES"
EXC = {"Hoja1","EDITAR (2)","Instrucciones",HOJA_MARCADORES,"Tabla 1","Banderas"}

def is_time(v): return isinstance(v, datetime.time)

def main():
    if not Path(EXCEL_PATH).exists():
        print("No encuentro", EXCEL_PATH); return
    wb  = openpyxl.load_workbook(EXCEL_PATH, data_only=True)   # valores calculados
    wbf = openpyxl.load_workbook(EXCEL_PATH, keep_vba=True)    # formulas
    mf  = wb[HOJA_MARCADORES]

    master = []
    for r in range(1, mf.max_row+1):
        b=mf.cell(r,2).value; c=mf.cell(r,3).value
        d=mf.cell(r,4).value; f=mf.cell(r,6).value
        if is_time(b) and d and f:
            dia = c.date().isoformat() if hasattr(c,"date") else str(c)[:10]
            master.append({"row":r,"h":b.strftime("%H:%M"),"d":dia,
                           "l":str(d).strip(),"v":str(f).strip()})

    players=[]
    for s in wb.sheetnames:
        if s in EXC: continue
        ws=wb[s]; wsf=wbf[s]
        name=str(ws["C2"].value or s).strip()
        preds=[]
        for r in range(1, ws.max_row+1):
            hf=wsf.cell(r,8).value      # formula H -> =+Marcadores_FINALES!$E$<fila>
            e=ws.cell(r,5).value; g=ws.cell(r,7).value
            if isinstance(hf,str) and "Marcadores_FINALES" in hf:
                m=re.search(r"\$E\$(\d+)", hf)
                if m and e is not None and g is not None:
                    preds.append({"mrow":int(m.group(1)),"pl":e,"pv":g})
        players.append({"n":name,"sheet":s,"preds":preds})

    Path("snapshot.json").write_text(
        json.dumps({"master":master,"players":players}, ensure_ascii=False),
        encoding="utf-8")
    print(f"snapshot.json creado: {len(master)} partidos, {len(players)} participantes.")

if __name__ == "__main__":
    main()
