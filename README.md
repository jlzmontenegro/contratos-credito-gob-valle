# Contratos de crédito público · Gobernación del Valle del Cauca (Hacienda) 2020–2026

Dashboard web (mobile first) para analizar los contratos de crédito público registrados en SECOP II, a partir de la pestaña **«ANDANDO»** del archivo `data/contratos.xlsx`.

## Qué incluye

- **Indicadores**: monto total, plazo promedio ponderado, vida remanente, saldo teórico, modificaciones, mayor contrato.
- **Plazo frente a monto** (análisis central): dispersión con tendencia lineal, correlaciones de Pearson y Spearman, R², pendiente, monto por rango de plazo, carga anual implícita (monto ÷ plazo) y plazo por acreedor. Incluye un texto de hallazgos que se recalcula con cada filtro.
- **Acreedores**: monto por entidad, participación por tipo de acreedor y concentración (índice HHI).
- **Evolución y vigencia**: monto firmado por año y cronograma de vigencia de cada contrato.
- **Vencimientos**: monto por año de vencimiento, amortización anual estimada y saldo teórico (supuesto lineal).
- **Relación de contratos**: número de contrato, acreedor, valor, objeto, fechas de inicio, fin y firma, plazo, estado y botón **Ver en SECOP**. Las columnas se pueden ordenar, filtrar (por ejemplo `>50` en Valor = más de 50 mil millones), mostrar u ocultar y reordenar arrastrando el encabezado. En el celular se ven como tarjetas o como tabla.
- **Filtros interactivos**: año de firma, acreedor, tipo de acreedor, administración, estado, plazo, monto y búsqueda libre. Al tocar las barras de las gráficas se filtra por ese acreedor, rango de plazo o año. Los filtros quedan en la URL, así que se puede compartir una vista filtrada.
- **Reportes**: **PDF** (indicadores, hallazgos, gráficas, relación de contratos con enlaces a SECOP y notas metodológicas) y **Excel** (resumen, contratos y una hoja por cada análisis), siempre con los filtros aplicados.
- **Calidad de datos**: se descartan los registros duplicados (sin fecha de inicio o con proveedor «cuenta no habilitada» y con el mismo valor y fecha de fin que un contrato válido). La sección de metodología señala referencias repetidas, contratos con fecha de fin cumplida y días adicionados.

## Actualizar los datos

1. Reemplace `data/contratos.xlsx` por la nueva versión del Excel (debe conservar la pestaña «ANDANDO» y los mismos encabezados).
2. Haga commit y push. GitHub Pages publica el cambio en uno o dos minutos.

También puede usar el botón **Cargar Excel** del dashboard para analizar otro archivo en el navegador sin publicarlo.

## Publicación en GitHub Pages

En el repositorio: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, elija la rama que contiene el sitio y la carpeta `/ (root)`, y guarde. El sitio quedará en `https://<usuario>.github.io/contratos-credito-gob-valle/`.

## Ejecutar localmente

El navegador no permite leer el Excel si abre `index.html` con doble clic, así que hay que servir la carpeta:

```bash
python3 -m http.server 8000
# abrir http://localhost:8000
```

## Estructura

```
index.html              página del dashboard
assets/styles.css       estilos (mobile first, tema claro/oscuro)
assets/app.js           lectura del Excel, cálculos, gráficas, tabla y reportes
assets/vendor/          Chart.js 4.4.1, SheetJS 0.18.5, jsPDF 2.5.1, jsPDF-AutoTable 3.8.2
data/contratos.xlsx     datos fuente (SECOP II)
```

## Notas metodológicas

- **Plazo**: años entre la fecha de inicio y la fecha de fin (días ÷ 365,25).
- **Carga anual implícita**: valor ÷ plazo, es decir, el capital que habría que amortizar cada año con pagos lineales.
- **Amortización y saldo teórico**: reparten el valor de forma lineal entre el inicio y el fin, sin periodos de gracia, desembolsos parciales ni intereses. Sirven para dimensionar el perfil de la deuda, pero no son el saldo real.
- **Administración**: periodo de gobierno de cuatro años según la fecha de firma.
- **Tipo de acreedor**: Findeter, Infivalle y Banco Agrario se agrupan como banca pública y de fomento; el resto, como banca privada.
