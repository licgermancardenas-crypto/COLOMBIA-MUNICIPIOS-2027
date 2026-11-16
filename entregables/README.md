# ATLAS COLOMBIA 2027 · Base maestra de inteligencia electoral

**Mapa interactivo en línea:** https://licgermancardenas-crypto.github.io/COLOMBIA-MUNICIPIOS-2027/
(acceso directo a un municipio: `.../entregables/MAPA_INTERACTIVO/ir/<código DIVIPOLA>.html`, por ejemplo `ir/05001.html`)

Elecciones territoriales del **31 de octubre de 2027** (período 2028-2031), según la Resolución 12338 de 2026 de la Registraduría.
Fecha de corte de los datos: **15-11-2026**.

## Entregables (`entregables/`)
| Archivo | Contenido |
|---|---|
| `BASE_MAESTRA_ATLAS_COLOMBIA_2027.xlsx` | 37 hojas: catálogos, FICHA por municipio, resultados, indicadores, simulador, modelo base, calidad, dashboard y 5 mapas interactivos |
| `DATA_DICTIONARY_ATLAS_COLOMBIA.xlsx` | Diccionario de columnas, fórmulas y función de cada hoja |
| `DATA_VALIDATION_REPORT.xlsx` | Controles de calidad, resumen y log del ETL |
| `DATABASE_SCHEMA.sql` | Esquema PostgreSQL + PostGIS por capas (raw/staging/clean/analytics/predictive/presentation). No se probó contra un servidor. |
| `CATALOGOS_DIVIPOLA.csv` | 1.122 entidades DIVIPOLA con población, pobreza y desempeño municipal |
| `CATALOGOS_POLITICOS.csv` | Partidos nacionales + coaliciones, GSC y organizaciones locales |
| `RESULTADOS_HISTORICOS.csv` | Alcaldía y gobernación 2003-2023, una fila por candidato |
| `MUNICIPIOS_COLOMBIA_MGN_simplificado.geojson` | Geometría municipal (MGN simplificada) con su código DIVIPOLA |
| `MAPA_INTERACTIVO/` | Mapa web: departamentos → municipios → sectores censales 2018 (y manzanas en vivo) |
| `ETL_PIPELINE/` | Copia de los scripts para reproducir todo |

## Cómo regenerar
```bash
bash etl/00_download.sh              # fuentes -> raw/
python etl/parse_wiki_2023.py raw    # Wikipedia 2023 -> raw/*_wiki.csv
python etl/01_build_data.py          # raw -> data/clean/*.csv (tablas normalizadas)
python etl/02_build_excel.py         # clean -> entregables/
python etl/03_maps_xml.py            # dibuja los mapas interactivos (no necesita Excel)
```
Requiere Python 3.12+ con pandas, geopandas, openpyxl y matplotlib. Los pasos se corren uno tras otro, porque la máquina tiene poca RAM.

## MAPA INTERACTIVO (web)
`entregables/MAPA_INTERACTIVO/ATLAS_MAPA_COLOMBIA_2027.html` se abre con doble clic en el navegador. Los archivos de la carpeta viajan juntos.

Navegación en tres niveles:
1. **Colombia:** clic en un departamento.
2. **Departamento:** gobernador y alcaldías por fuerza; clic en un municipio.
3. **Municipio:** ficha completa más sus **sectores censales DANE 2018** (12.073 en todo el país), coloreados por la variable que elijas: población, densidad, % mayores de 60, % jóvenes, % con educación superior, % indígena, % afro, % migrantes, analfabetismo, estrato predominante, % viviendas sin acueducto o % viviendas con internet.

Al hacer clic en un sector se ve su detalle. Con zoom alto se pueden activar las **manzanas por estrato**, que se consultan en vivo.

**Puestos de votación:** son 13.490, de las legislativas del 8-mar-2026 (escrutinio de Senado y Cámara, Esri Colombia con datos de la Registraduría).
- Se dibujan como círculos sobre los sectores censales. El tamaño es el potencial electoral y el color es la lista ganadora en Senado o en Cámara, o la participación.
- Al hacer clic en un puesto se ve su ficha: dirección, comuna o localidad, potencial por sexo, participación y ganador en cada cámara.
- El panel del municipio resume los puestos y muestra un ranking de los de mayor potencial; al hacer clic en uno, el mapa vuela hasta él.
- Los «Puesto Censo» son puestos administrativos de la Registraduría para cédulas sin sede asignada. Se muestran en blanco y quedan fuera del ranking.

Hay filtros por **región** (Caribe, Andina, Pacífica, Orinoquía, Amazonía, Insular), por **departamento** y por **fuerza política** (la del alcalde electo en 2023; se combina con los otros dos y muestra cuántas alcaldías y qué población gobierna esa fuerza en el ámbito elegido, por departamento y con la lista de municipios), buscador de municipios y temas de color (fuerza, swing, bastiones, IPM). La región resalta sus departamentos, hace zoom y muestra un resumen con gobernaciones, alcaldías por fuerza y competitividad. En el Excel, MUNICIPIOS, DEPARTAMENTOS y ALCALDES traen la columna `region` para filtrar. Desde el Excel, las hojas MAPA y la FICHA abren el mapa directamente en ese municipio o departamento. Lo hacen a través de los accesos directos de `MAPA_INTERACTIVO/ir/` (un archivo por municipio y por departamento), así que no dependen de que Excel le pase el municipio al navegador.

Funciona sin internet, salvo el mapa base y las manzanas. Las fuentes de los sectores y las manzanas son los servicios Esri Colombia LivingAtlas (CNPV 2018), porque el geoportal del DANE bloquea las conexiones desde fuera de Colombia.

Para regenerarlo: `python etl/05_download_sectores.py`, `python etl/06_download_puestos.py` y después `python etl/04_build_webmap.py`.

## FICHA
En la celda amarilla se elige un municipio, con el formato `NOMBRE · DEPARTAMENTO`. La ficha arma todo con fórmulas sobre las tablas de la base:
- datos generales y sociales;
- alcalde, organización, coalición y antecedentes;
- gobernador del departamento;
- clasificación de swing, bastión y probabilidad del modelo;
- historia de la alcaldía 2003-2023 y resultados 2019 por candidato;
- actores 2027 del municipio y del departamento;
- enlaces y un mini mapa que resalta el municipio elegido.

La hoja auxiliar GEO FICHA guarda las coordenadas del mini mapa.

## Mapas
Hay cinco hojas MAPA: POLÍTICO, GOBERNACIONES, SWING, BASTIONES y TERRITORIO. En cada una, todos los municipios (o departamentos) son formas con su contorno real:
- al pasar el mouse se ven sus datos;
- con un clic se va a su fila en MUNICIPIOS o DEPARTAMENTOS.

El recuadro de la derecha amplía la zona andina, y arriba a la izquierda están San Andrés y Providencia. Las formas se escriben directo en el XML del .xlsx porque el Excel de esta máquina está en modo de funcionalidad reducida (licencia vencida) y no permite editar por automatización.

## Qué hay y de dónde sale
| Bloque | Cobertura | Fuente |
|---|---|---|
| Territorio | 1.103 municipios + 18 áreas no municipalizadas + San Andrés | DANE DIVIPOLA |
| Alcaldía 2003-2019 | Todos los candidatos, blancos, nulos, no marcados y censo (desde 2007) | CEDAE – Registraduría (oficial consolidado) |
| Gobernación 2003-2019 | Todos los candidatos, por municipio | CEDAE |
| Alcaldía 2023 | Ganador de 1.079 de 1.103 municipios | Servicio Esri con datos de la Registraduría; 48 municipios repuestos desde Wikipedia |
| Gobernación 2023 | Todos los candidatos y totales, por departamento | Wikipedia (hay que contrastarla con el escrutinio) |
| Población 2018/2026/2027, NBI e IPM 2018, MDM 2020 | Municipal | DANE, DNP |
| Partidos con personería 2026 | 21 organizaciones | El Espectador, 05-08-2026 (verificar con el CNE) |
| Aspirantes y prospectos 2027 | 40 menciones, de las cuales 3 son aspirantes o precandidatos | Prensa (Infobae, La FM, El Nuevo Siglo, El Colombiano) |

## Limitaciones conocidas (detalle en la hoja CALIDAD DATOS)
1. **Alcaldía 2023 sólo trae el ganador confiable.** En la fuente Esri, el 2.º y el 3.º promedian 15 % de los votos (en 2019 el 2.º promediaba 34 %), y los tres registros juntos suman alrededor del 66 % de los válidos. Por eso no se calculan márgenes ni NEP para 2023, y SWING y SIMULADOR trabajan sobre 2019, el último resultado oficial completo. El escrutinio oficial 2023 queda pendiente: los portales de la Registraduría bloquearon la descarga automática.
2. **Nombres de partido 2003-2019.** La fuente CEDAE trae sólo el código de la organización. El nombre se infirió del archivo de concejo del mismo año y del servicio Esri 2019. Los que no se pudieron resolver figuran como «SIN NOMBRE EN FUENTE (código X)»: 67 % de las candidaturas en 2003 y 51 % en 2007; menos del 15 % desde 2011.
3. **24 municipios sin alcalde 2023 identificado**, en su mayoría de Santander, además de Nuevo Belén de Bajirá, que es un municipio nuevo.
4. **No se cargaron** concejos, asambleas, JAL, resultados mesa a mesa (ver `etl/03_mesa_a_mesa.md`), encuestas, TerriData/FUT/CUIPO ni la categoría municipal (Ley 617). Las estructuras ya están creadas.
5. **Las coaliciones no se atribuyen a partidos.** Los bastiones y la alternancia se miden sólo para partidos canónicos; una coalición, un GSC o una organización local queda como «indeterminada».
6. **El `candidato_id` enlaza candidaturas entre años por coincidencia de nombre.** Puede unir homónimos o no reconocer variantes de un mismo nombre.
7. **El modelo base v0.1** estima la probabilidad de que la misma fuerza retenga la alcaldía según su racha histórica. En el backtest 2019→2023 el Brier es 0,1066, frente a 0,1089 de una constante: sirve como línea de base, no como pronóstico de candidatos.
8. **La condición actual de cada autoridad está por verificar.** Puede haber cambios por elecciones atípicas, nulidades o renuncias; por ejemplo, la elección del gobernador de San Andrés fue anulada según la prensa.

## Principios aplicados
No se inventan datos. Partidos ≠ coaliciones, alcaldes ≠ candidatos. Alcaldía y gobernación van en tablas separadas. Cada fila tiene `fuente_id`. Los pronósticos se versionan con `forecast_id` y nunca se sobrescriben. Ningún resultado 2023 se presenta como escrutinio definitivo.
