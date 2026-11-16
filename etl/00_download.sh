#!/usr/bin/env bash
# ATLAS COLOMBIA 2027 – Paso 0: descarga de fuentes a raw/ (reproducible).
# Uso: bash etl/00_download.sh   (desde la raíz del proyecto)
set -euo pipefail
cd "$(dirname "$0")/../raw"

# DIVIPOLA (DANE, datos.gov.co)
curl -s "https://www.datos.gov.co/resource/gdxc-w37w.json?\$limit=5000" -o divipola.json

# Resultados históricos CEDAE – Registraduría (alcaldía, gobernación; concejo sólo para nombres de partido)
B=https://s3.amazonaws.com/cedae.dskt.ch/resultados
for y in 2003 2007 2011 2015 2019; do
  curl -s "$B/Alcaldia/${y}_alcaldia.dta.csv"       -o ${y}_alcaldia.csv
  curl -s "$B/Gobernacion/${y}_gobernacion.dta.csv" -o ${y}_gobernacion.csv
  curl -s "$B/Concejo/${y}_concejo.dta.csv"         -o ${y}_concejo.csv
done

# Alcaldía 2019 y 2023 (servicio Esri Colombia con datos de la Registraduría) + geometría municipal MGN
E=https://ags.esri.co/arcgis/rest/services/Elecciones_Colombia
for y in 2019 2023; do
  curl -s "$E/Resultados_Electorales_Alcaldia_Municipio_Colombia_$y/MapServer/0/query?where=1%3D1&outFields=*&returnGeometry=false&f=json&resultRecordCount=2000" -o esri_alc_$y.json
done
curl -s "$E/Resultados_Electorales_Alcaldia_Municipio_Colombia_2023/MapServer/0/query?where=1%3D1&outFields=codigoMunicipalMGNCompleto,nombreMunicipio&returnGeometry=true&outSR=4326&maxAllowableOffset=0.004&geometryPrecision=4&f=geojson" -o mun_2023.geojson

# DANE: proyecciones de población, NBI e IPM
curl -s "https://www.dane.gov.co/files/censo2018/proyecciones-de-poblacion/Municipal/PPED-AreaMun-2018-2042_VP.xlsx" -o dane_proy_area.xlsx
curl -s "https://www.dane.gov.co/files/censo2018/informacion-tecnica/CNPV-2018-NBI-DIVIPOLA-2021.xlsx" -o dane_nbi_2021.xlsx
curl -s "https://www.dane.gov.co/files/investigaciones/condiciones_vida/pobreza/2020/Anexos_medida_pobreza_mun_CNPV_2018.xlsx" -o dane_ipm_mun.xlsx

# DNP: Medición de Desempeño Municipal 2020
curl -s -G "https://www.datos.gov.co/resource/nkjx-rsq7.json" \
  --data-urlencode "\$select=codigo_entidad,indicador,dato,etiqueta" \
  --data-urlencode "\$where=anio='2020' AND codigo_unico in('080010001','080010008','080010002','080010005')" \
  --data-urlencode "\$limit=10000" -o mdm2020.json

# Wikipedia "Elecciones regionales de Colombia de 2023": se guardó como markdown en raw/wiki_reg2023.md
# (descargado con Firecrawl). Luego: python etl/parse_wiki_2023.py raw
echo "Descargas completas."
