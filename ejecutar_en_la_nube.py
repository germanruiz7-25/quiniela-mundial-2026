#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
  ejecutar_en_la_nube.py
  --------------------------------------------------------------------------
  Este archivo es el que corre GitHub Actions (la nube) en automatico.
  Hace, en orden:
     1. Toma la "foto" del Excel (snapshot) ANTES de tocar nada -> asi
        agregar o modificar participantes/vaticinios se refleja solo.
     2. Consulta la API y escribe los marcadores nuevos (filtro de calidad).
     3. Recalcula puntos y regenera el dashboard publico.

  Nota tecnica importante: openpyxl no recalcula formulas, y al guardar el
  .xlsm pierde los nombres calculados. Por eso tomamos el snapshot UNA sola
  vez al inicio (cuando el Excel aun tiene sus valores) y lo reutilizamos.
  Los marcadores (E/G) son valores fijos, asi que sobreviven el guardado y
  se leen bien para recalcular puntos.

  NO toma captura ni abre WhatsApp: eso no se puede en la nube. El enlace
  publico del dashboard se actualiza solo; desde ahi mandas la imagen al
  WhatsApp con un screenshot normal cuando quieras.
"""
import construir_snapshot
import actualizar_quiniela as A

def run():
    print(">> 1/3  Tomando foto del Excel (snapshot) ANTES de modificar...")
    construir_snapshot.main()              # crea snapshot.json desde el Excel intacto
    snap = A.load_snapshot()               # lo cargamos en memoria y NO lo volvemos a tocar

    print(">> 2/3  Consultando resultados y actualizando marcadores...")
    results = A.fetch_api_results()
    applied = A.update_excel(results, snap) if results else 0
    print(f"   {applied} marcador(es) nuevo(s).")

    print(">> 3/3  Recalculando puntos y regenerando dashboard...")
    # Usamos el MISMO snapshot de memoria (nombres intactos) +
    # los marcadores E/G ya guardados (que si sobreviven). Asi el dashboard
    # queda perfecto sin depender de las formulas del Excel.
    A.regenerate_dashboard(snap)
    print("Listo. Dashboard actualizado.")

if __name__ == "__main__":
    run()
