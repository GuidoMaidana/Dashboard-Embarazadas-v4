Dashboard de Salud Materna — Ciudad de Corrientes
=================================================

Grupo: Atención médica de embarazadas y cobertura de centros de salud
Integrantes: Daniel Dure · Maidana Guido

Despliegue
----------
Subir index.html, data.js y vercel.json al repositorio de GitHub conectado al proyecto
de Vercel. No requiere build ni dependencias: es HTML estático. La carpeta generacion/
contiene los insumos y scripts de limpieza; no es necesaria para el funcionamiento del
sitio.
Para probar en local:  python3 -m http.server 8000  →  http://localhost:8000

Archivos
--------
index.html   Toda la interfaz, los estilos y la lógica.
data.js      Datos ya procesados (consultas, SAPS, CAPS, hospitales, centros,
             áreas) en formato legible: window.DASHBOARD_DATA, un campo por línea.
vercel.json  Configuración de despliegue.
generacion/  Insumos y scripts reproducibles (no son necesarios para el sitio):
               datos_salud_optimizados.csv  base limpia: 207 filas de consulta + 99 establecimientos
               zonas.geojson                polígonos de las áreas programáticas I/II/III
               generar_data_js.py           genera data.js desde el CSV y las zonas
               generar_datos_limpios.py     genera el CSV desde las planillas (requiere la carpeta
                                            "Excel datos" junto al script)
               LIMPIEZA_DATOS.md            informe de limpieza: criterios, alias, duplicados, cifras

Versión 4 — 19/10/2026 (datos reprocesados)
-------------------------------------------
data.js se regenera desde la base limpia (antes tenía los datos crudos incrustados):
- 1.455 → 1.420 consultas: se quitaron 6 filas duplicadas exactas y 1 fila vacía (214 → 207 filas).
- 31 → 30 SAPS: «DR. BLUGERMAN» y «DR. PEDRO BLUGERMAN» quedan unificados desde la base
  (se eliminó el alias del código).
- Catálogo de SAPS completo: 37 establecimientos (la versión anterior usaba una planilla de 34).
- Numeral corregido: la fuente repetía «XXXIII» en Sussini; por secuencia corresponde a SAP XXXII.
- Textos normalizados: tildes, mayúsculas, rangos («16-20 años»), meses y etapas de pandemia.
- Área programática (I/II/III) asignada a los 99 establecimientos por punto-en-polígono.
- Cobertura completa: los 30 SAPS con consultas tienen coordenadas (antes: 6 de 31 sin georreferenciar).
- Cruce de nombres directo (mismos nombres canónicos en consultas y establecimientos): se eliminó
  la similitud de bigramas y la corrección de transliteraciones del código.
- data.js quedó en varias líneas (legible y editable a mano) y se agregó la carpeta generacion/
  con los scripts y la base que lo generan.
- Los gráficos usan un nombre corto de SAPS (sin numeral ni barrio) para que las etiquetas entren
  bien; la tabla, los popups y la exportación usan el nombre completo.
- La exportación CSV de la selección ahora incluye la columna Zona.
- Metodología y proyecto: fuentes y fecha de descarga, transformaciones, cálculos, supuestos
  y límites.
- Mapa: se reemplazaron las teselas de CARTO por cuatro capas base sin clave de API (Esri calles,
  Esri satélite, OSM France y OpenTopoMap), con selector «Mapa base» y paso automático a la
  siguiente capa si una no logra cargar mosaicos.
- Mapa: las capas base cargan teselas en todo el recuadro visible (incluido el estado
  «Centrar»), sin franjas vacías en los bordes.
- Paleta visual: fondos y bordes grises, acento verde institucional y gráficos con
  colores variados para distinguir mejor las series. La tarjeta de alerta de caída
  mar–jun lleva borde naranja.
- Mapa: SAPS en verde; CAPS en azul, hospitales en rojo, centros privados en naranja
  y áreas programáticas en gris, para diferenciarlos de un vistazo.
- Se redujo el redondeo de las esquinas de paneles, tarjetas, controles y tablas.
- «Metodología» y los datos del proyecto (objetivos, fuentes e integrantes) se unificaron
  en una sola sección.
- «Fuentes de datos» y la fecha de descarga (29/08/2026) quedaron unificadas en
  «Metodología y proyecto».
- Se quitaron los recuadros de «Cobertura del dato» (Consultas registradas, SAPS con
  datos, SAPS sin coordenadas y Edad «Sin dato»): repetían totales ya visibles o
  mostraban valores en cero.
- Se quitó la sección «Reproducibilidad» y las referencias técnicas a data.js y
  generacion/ dentro del dashboard.
- Se simplificaron el estado superior y el pie: ya no muestran cantidad de registros ni
  fecha de reprocesamiento.
- El contexto del ASPO del gráfico de evolución quedó como tooltip del botón «Contexto».
- «Detalle de consultas» se movió debajo del mapa y ahora es un acordeón cerrado por
  defecto, pensado para consultas específicas.

Versión 3 — 23/09/2026 (devolución del 21/09)
---------------------------------------------
- Nombres normalizados en pantalla (mayúsculas + DR. BLUGERMAN = DR. PEDRO BLUGERMAN): 31 → 30 SAPS.
- Gráfico de evolución con banda ASPO (desde 20/03/2020) y línea de SAPS que informaron por mes.
- Panel «Cobertura del dato» y panel «Lectura del período» (caída ASPO, recuperación,
  SAPS que informan, participación adolescente).
- Distribución por edad en barras; heatmap SAPS × mes con meses activos, media y desvío.
- Clustering (Leaflet.markercluster) en CAPS, hospitales y centros privados.
- Sección «Metodología» (fuente, transformaciones, cálculos, supuestos, límites, reproducibilidad).
  (En la v4 se completaron la fecha de descarga y el enlace a los scripts que generan data.js.)

Versiones anteriores (hasta el 21/09/2026)
------------------------------------------
- Modo oscuro automático (prefers-color-scheme) con botón Automático → Claro → Oscuro y
  preferencia guardada en localStorage; se aplica antes del primer pintado.
- Mapa acotado al ejido de la capital (maxBounds, zoom mínimo 11, teselas sin repetición)
  y sin descarga del polígono provincial externo.
- Pie de página con integrantes, grupo, objetivos y fuentes.
- Fallos corregidos: (a) el gráfico de edad que nunca se dibujaba (choque entre la variable
  «ages» y el id del canvas), que además cortaba el resto del render; (b) variables globales
  implícitas → todo encapsulado en un IIFE con 'use strict'; (c) botón «Ver Capital» →
  «Centrar» sobre la caja fija; (d) capas excluyentes → capas combinables con aria-pressed;
  (e) promedio por registro → promedio mensual; (f) registros sin rango etario («Sin dato»:
  la v4 los elimina en la limpieza); (g) la rueda del mouse ya no secuestra el scroll;
  (h) transliteraciones ANIOS→AÑOS y Caniada→Cañada (la v4 las corrige en la base);
  (i) fechas «ene 2020» en vez de «01/2020»; (j) si Leaflet o Chart.js no cargan, la página
  ya no se rompe entera.
- Agregados: barras apiladas por rango etario; KPI de establecimientos georreferenciados;
  puntos de SAPS que cambian de tamaño según el filtro; exportación CSV de la selección;
  tooltips, accesibilidad y mejoras responsive.
- Nota histórica: la coincidencia de nombres (hoy directa) se resolvía con similitud de
  bigramas, y 6 de los 31 SAPS de consultas no tenían coordenadas publicadas (resuelto: 0 de 30).

Fuentes de datos
----------------
https://datos.ciudaddecorrientes.gov.ar/dataset/atencion_embarazadas
https://datos.ciudaddecorrientes.gov.ar/dataset/centros-de-salud

Cartografía base: Esri (calles y satélite), OpenStreetMap y OpenTopoMap, sin claves
de API. El mapa permite elegir entre esas cuatro capas y pasa solo a la siguiente si una
no logra cargar mosaicos.
Áreas programáticas: polígonos oficiales y datos de establecimientos incluidos en las
planillas de origen (ver generacion/LIMPIEZA_DATOS.md).
