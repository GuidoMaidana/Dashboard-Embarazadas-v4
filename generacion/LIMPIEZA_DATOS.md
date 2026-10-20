# Informe de limpieza y optimización de datos

**Proyecto:** Dashboard de salud materna — consultas de embarazadas por SAPS (Corrientes Capital, ene–sep 2020)
**Fecha del informe:** octubre 2026
**Archivo generado:** `datos_salud_optimizados.csv` (raíz del proyecto, fuera de `Excel datos/`)
**Script reproducible:** `generar_datos_limpios.py`

---

## 1. Objetivo

Consolidar en una única tabla limpia y trazable toda la información disponible en la carpeta
`Excel datos/`: las **consultas mensuales de embarazadas** (enero–septiembre 2020) y el
**catálogo de establecimientos de salud** (SAPS, CAPS, hospitales públicos y centros privados),
corrigiendo errores de tipeo, acentos, mayúsculas, duplicados y nombres inconsistentes, e
incorporando la zona sanitaria (Área I, II o III) a cada establecimiento.

---

## 2. Fuentes utilizadas y descartadas

| Archivo | Contenido | Decisión |
|---|---|---|
| `Consultas_embarazadas.xlsx` | 214 filas de consultas mensuales por SAPS y grupo etario | **Usada** (fuente principal de consultas) |
| `datos_salud_combinados.xlsx` | 37 SAPS con nombre, barrio, dirección y coordenadas + polígonos de las 3 zonas | **Usada** (catálogo SAPS autoritativo, 37 registros completos) |
| `Saps.xlsx` | 34 SAPS | **Descartada**: incompleta — faltaban Casusso, Rossi Candia y Santa Marta respecto de la fuente combinada |
| `Caps.xlsx` | 14 CAPS | Usada |
| `Hospitales.xlsx` | 8 hospitales públicos | Usada |
| `centros de salud.xlsx` | 40 centros privados | Usada |
| `centros de salud (1).xlsx` | Copia idéntica de la anterior (verificado por SHA256) | **Descartada** por duplicada |

**Resultado:** 99 establecimientos (37 SAPS + 14 CAPS + 8 hospitales + 40 privados).

---

## 3. Limpieza de la base de consultas

### 3.1 Filas eliminadas

| Problema | Filas | Detalle |
|---|---|---|
| Fila vacía (cantidad y grupo etario `NULL`) | 1 | `DR. SEMPER`, enero 2020 |
| Filas duplicadas exactas (misma SAPS, fecha, grupo etario y cantidad) | 6 | Se conservó **un solo** registro de cada par; no se sumaron (ver detalle abajo) |

Detalle de los duplicados exactos (se removió 1 fila de cada par, 35 consultas que
inflaban el total):

| SAPS | Mes | Grupo etario | Cantidad |
|---|---|---|---|
| Dr. Chercoff | 01-2020 | 21-39 años | 1 |
| Dr. Calmanash | 01-2020 | 16-20 años | 10 |
| Dr. Calmanash | 01-2020 | 21-39 años | 18 |
| Gdor. Elías Galván | 05-2020 | 12-15 años | 1 |
| Gdor. Elías Galván | 05-2020 | 21-39 años | 1 |
| Villa Chiquita | 07-2020 | 21-39 años | 4 |

> Total de consultas corregido: **1.455 (crudo) → 1.420 (limpio)**.

### 3.2 Normalización de valores

- **Grupos etarios:** `12 - 15 ANIOS` → `12-15 años` (edades 12/15), `16 - 20 ANIOS` → `16-20 años`
  (16/20), `21 - 39 ANIOS` → `21-39 años` (21/39), `>=40 ANIOS` → `40 años o más` (40, sin tope).
- **Fechas:** convertidas a periodo `AAAA-MM` con nombre de mes en español (`Enero`…`Septiembre`).
- **Etapas de pandemia:** Enero–Febrero = `Pre-ASPO`; Marzo–Junio = `ASPO`; Julio–Septiembre = `Recuperación`.
- **Servicio:** unificado como `Embarazadas del Área Programática`.

---

## 4. Normalización de nombres SAPS (31 → 30) — observación de la coordinadora

La base de consultas usa 31 nombres distintos que, tras el mapeo, corresponden a **30 SAPS**
(el problema señalado en la revisión). Mapeo de alias aplicado:

| Nombre en consultas | SAPS destino (canónico) |
|---|---|
| `DR. BLUGERMAN` **y** `DR. PEDRO BLUGERMAN` | **SAP XXXIII - Dr. Pedro Blugerman** (unificados: 29 + 185 = 214 consultas) |
| `DR. PIRCHI` | SAP III - Dr. Oscar E. Pirchi (Bº Anahí) — Cristo Obrero 310 |
| `DR. BENJAMIN SERRANO` | SAP XVII - Dr. Benjamín Serrano (Bº Parque Cadenas Norte) — Ruta 5 km 2,5 |
| `DR. FLIER` | SAPS XXVII - Dr. Juan Carlos Benítez Suárez (SAPS Lomas) |
| `DR. MANUEL CASSUSO` | SAP IX - Dr. Manuel Casusso |
| `DR. ROSSI CANDIA` | SAP VIII - Dr. Miguel Rossi Candia |
| `CICHERO` | SAP XXXV - Dr. Romilio Monzón (Bº Cichero) |
| `DR MANUEL A. GONZALEZ` | SAP XIV - Dr. Manuel A. González |
| `GDOR. ELIAS GALVAN` | SAPS XXXIV - Dr. Elías Galván |
| `ITATI` | SAP IV - Bº Itatí |
| `QUINTA FERRE` | SAP V - Bº Quinta Ferré |
| `RIO PARANA` | SAP XXXI - Bº Río Paraná |
| `SANTA MARTA` | SAPS XXIX - Bº Santa Marta |
| `SANTA MARGARITA` | SAP XIII - Bº Santa Margarita |
| `SAPUCAY` | SAPS XVIII - Bº Sapucay |
| `VILLA CHIQUITA` | SAP XI - Bº Villa Chiquita |
| `ESPERANZA` | SAP XII - Bº Esperanza |
| `GÜEMES` | SAP XXXVII - Bº Güemes |
| `PRIMERA JUNTA` | SAP XIX - Bº Primera Junta |
| `SAN MARCOS` | SAP VII - Bº San Marcos |
| `DR. ANIBAL MALVIDO` | SAP XX - Dr. Aníbal Malvido |
| `DR. ALBERTO LIFSCHITZ` | SAP XXV - Dr. Alberto Lifschitz |
| `DR. SANTIAGO LORENZO` | SAP XXVI - Dr. Santiago Lorenzo |
| `DR. GUILLERMO RAWSON` | SAP I - Dr. Guillermo Rawson |
| `DR. SEMPER` | SAP XV - Dr. Juan Carlos Semper |
| `DR. CHERCOFF` | SAP XXIII - Dr. Juan Chercoff |
| `DR. SUSSINI` | SAP XXXII - Dr. Miguel Sussini |
| `DR. CALMANASH` | SAP XXX - Dr. Saúl Calmanash |
| `DR. KORIMBLUM` | SAP X - Dr. Tomás Korimblum |
| `DR. MAURICIO OPEN` | SAPS XXIV - Dr. Mauricio Open |

**Corrección adicional en el catálogo:** la fuente `datos_salud_combinados.xlsx` repetía el numeral
`XXXIII` para dos establecimientos distintos (Blugerman y Sussini). Por secuencia y gid, se corrigió
el de Sussini a **SAP XXXII - Dr. Miguel Sussini**.

---

## 5. Coordenadas y zonas sanitarias

- Se conservaron las coordenadas de la fuente combinada (redondeadas a 7 decimales).
- A cada uno de los 99 establecimientos se le asignó su **zona** (Área I, II o III) por
  *point-in-polygon* contra los polígonos oficiales incluidos en `datos_salud_combinados.xlsx`.

| Zona | Establecimientos |
|---|---|
| Área I | 57 |
| Área II | 20 |
| Área III | 22 |

- **Los 30 SAPS con consultas tienen coordenadas (0 sin coordenadas)**, lo que responde al punto
  señalado por la coordinadora ("6 de 31 SAPS sin coordenadas"). Los 7 SAPS sin consultas en
  ene–sep 2020 también tienen coordenadas y zona: Cañada Quiroz, Pujol, Villa García,
  De Los Reyes Vidal, Centro de Vacunación, Paso Martínez y Santa Catalina.

---

## 6. Catálogo de centros (nombres y direcciones)

- **CAPS y hospitales:** se corrigieron mayúsculas y acentos (p. ej. `CAPS VI - Dr. Tiburcio
  Foncesa` → `Fonseca`, `HOSPITAL PSIQUIATRICO` → `Hospital Psiquiátrico`).
- **Centros privados:** 40 registros con nombre y dirección normalizados (tildes, abreviaturas
  `Av.`, `Gral.`, `N°`, `S/N`, `Mza.`, etc.).
- **Clínica del Niño:** aparecía dos veces con direcciones contiguas (Belgrano 1774 y 1760,
  24 m de distancia). Se conservaron **ambos** registros —podrían ser dos sedes del mismo
  centro— diferenciados como `Clínica del Niño (dirección A)` y `(dirección B)`, con una
  observación que advierte el posible duplicado.

---

## 7. Respuesta a las observaciones de la coordinadora

| Observación | Respuesta |
|---|---|
| Nombres SAPS inconsistentes / duplicados (`DR. BLUGERMAN` vs `DR. PEDRO BLUGERMAN`) | Resuelto: alias unificados → **30 SAPS** únicos (sección 4) |
| 6 de 31 SAPS sin coordenadas | Resuelto: los 30 SAPS con consultas tienen coordenadas; el mapeo de alias asignó la geometría correcta a cada uno |
| Total 1.455 consultas | **1.420** tras remover 6 filas duplicadas exactas (35 consultas) y 1 fila vacía |
| Caída de marzo (ASPO) sin interpretar | Se incorporó la columna `etapa_pandemia` (Pre-ASPO / ASPO / Recuperación) para permitir el análisis: marzo cae a 61 consultas (−83% vs. febrero) |
| Rango "Sin dato" | La única fila con "Sin dato"/NULL (DR. Semper, enero) se eliminó y quedó documentada |
| Falta de indicadores territoriales/poblacionales | Se agregó la **zona sanitaria** por establecimiento; ver limitación en la sección 9 (no se dispone de datos de población) |

---

## 8. Cifras clave post-limpieza

**Consultas por mes (total: 1.420):**

| Mes | Consultas | Etapa |
|---|---|---|
| Enero | 415 | Pre-ASPO |
| Febrero | 350 | Pre-ASPO |
| Marzo | 61 | ASPO |
| Abril | 99 | ASPO |
| Mayo | 53 | ASPO |
| Junio | 96 | ASPO |
| Julio | 221 | Recuperación |
| Agosto | 65 | Recuperación |
| Septiembre | 60 | Recuperación |

- Promedio mensual Pre-ASPO (ene–feb): **382,5**; ASPO (mar–jun): **77,3** (−79,8%);
  Recuperación (jul–sep): **115,3**.
- Adolescentes (12–20 años): **292 consultas (20,6% del total)**.

**Consultas por grupo etario:**

| Grupo | Consultas | % |
|---|---|---|
| 12-15 años | 22 | 1,5% |
| 16-20 años | 270 | 19,0% |
| 21-39 años | 1.077 | 75,8% |
| 40 años o más | 51 | 3,6% |

**SAPS con más consultas:** Blugerman 214 (8 meses), Rossi Candia 176, Calmanash 134,
Sussini 125, Malvido 115, Pirchi 90, Monzón (Cichero) 75, Itatí 57.

---

## 9. Formato del entregable y limitaciones

**Formato del CSV:** delimitador `;`, codificación UTF-8 con BOM, fin de línea CRLF,
campos con texto que contiene `;` entrecomillados (compatible con Excel en español).
Cada fila tiene una columna `tipo_registro` (`consulta` o `centro`):

- **207 filas `consulta`**: detalle mensual por SAPS y grupo etario (1.420 consultas).
- **99 filas `centro`**: catálogo completo con zona, coordenadas y totales anuales
  (`consultas_totales_2020`, `meses_con_registros_2020` solo se completan en filas de centro
  para evitar doble conteo).

**Limitaciones:**

- No se dispone de datos de población por barrio/zona, por lo que **no es posible calcular**
  consultas por 1.000 mujeres embarazadas ni ratios de cobertura (punto pendiente de la revisión).
- La base de consultas registra cantidades agregadas por mes y rango etario: no hay datos
  por persona, por lo que no pueden calcularse tasas individuales ni seguimiento.
- Para los SAPS sin registros en ene–sep 2020 no puede distinguirse entre "sin demanda",
  "sin registro" o "no informado"; se los conserva en el catálogo con total 0 y observación.
- Las coordenadas provienen de la fuente original; se redondearon a 7 decimales sin
  modificar su ubicación.

---

*Generado automáticamente a partir de `generar_datos_limpios.py`. Para reproducir el CSV:
`python generar_datos_limpios.py`.*
