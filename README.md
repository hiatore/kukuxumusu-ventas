# KUKUXUMUSU - Dashboard Ejecutivo de Ventas

Dashboard profesional para la gestión y análisis de ventas de las tiendas Kukuxumusu.

## Características

- **KPIs globales**: facturación total, ingresos periodo, unidades, stock, devoluciones
- **Tarjetas por tienda**: acumulado anual + mes en curso con producto estrella
- **Rankings filtrables**: por unidades o ingresos, con buscador en vivo
- **6 gráficos Chart.js**: facturación, tendencia mensual, categorías, ventas diarias
- **Modal detalle**: click en tienda → top 10 productos (unidades, ingresos, histórico) + categorías
- **Datos históricos** desde los ficheros individuales de cada tienda

## Cómo actualizar los datos

1. Descarga los Excel desde StockAgile a `~/Downloads/`
2. Ejecuta `python3 generar_dashboard.py`
3. Abre `dashboard_ejecutivo.html` en el navegador

## Tecnologías

- Python + Pandas (procesamiento de datos)
- Chart.js (gráficos)
- HTML + CSS + JS vanilla (frontend)
- GitHub + Render (hosting)

## Estructura

```
/
├── dashboard_ejecutivo.html   # Dashboard listo para usar
├── generar_dashboard.py       # Generador de datos
├── actualizar_dashboard.bat   # Atajo para Windows
├── servidor_dashboard.py      # Servidor web local
├── DASHBOARD_GLOBAL_TIENDAS.xlsx
├── analisis_rotacion_*.xlsx   # Datos por tienda
└── README.md
```
