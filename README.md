#  excel-report-automation

**Automatización de un reporte mensual recurrente en Excel con Python**: de un export "sucio" de ERP a un reporte formateado con KPIs, tablas dinámicas y gráfico, en segundos.

> ⚠️ **Nota:** inspirado en tareas que automaticé en mis prácticas (finanzas en Elanco Animal Health y administración en una consultora). Esas automatizaciones las hice en **Excel avanzado**. Este repositorio replica la idea en Python, con **datos 100% simulados**.

## Contexto real
- Automaticé reportes recurrentes con Excel avanzado y **reduje en un 70% el tiempo** de elaboración de tareas recurrentes.
- Analicé y registré datos contables para apoyar la gestión financiera diaria.

## Qué hace
1. **Simula un export de ERP** con problemas típicos: filas duplicadas, montos como texto (`"1.234.500"`), áreas mal escritas.
2. **Limpia** los datos: elimina duplicados, normaliza nombres y convierte montos.
3. **Calcula KPIs**: gasto total, documentos, ticket promedio, gasto vs presupuesto por área y alertas de desviación > 5%.
4. **Genera un Excel formateado** con 5 hojas (Resumen, Por área, Área × categoría, Semanal, Detalle), encabezados con estilo, formato de miles, paneles fijos y un gráfico de barras.

## Ejecutar
```bash
pip install -r requirements.txt
python report_automation.py --month 2025-06
# -> reporte_2025-06.xlsx
```

## Ejemplo de salida (simulado)
```
Filas crudas: 825 -> limpias: 800 (duplicados eliminados: 25)
       area   desviacion_%            estado
  Comercial           10.6 Sobre presupuesto
   Finanzas           -5.8                OK
     RR.HH.            6.0 Sobre presupuesto
Reporte generado: reporte_2025-06.xlsx en 0.2 s
```

## Conceptos aplicados
Automatización de procesos · limpieza de datos · control presupuestario · reporting · Python (pandas, openpyxl) · Excel

---
Autor: **Liam Esquivel González** · [LinkedIn](https://www.linkedin.com/in/liamesquivelg)
