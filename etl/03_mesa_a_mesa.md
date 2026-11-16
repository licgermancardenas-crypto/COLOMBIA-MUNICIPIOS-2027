# Carga mesa a mesa (pendiente)

Esta versión no trae datos mesa a mesa. La hoja `RESULTADOS MESA A MESA` y la tabla
`clean.fact_resultados_mesa` de `DATABASE_SCHEMA.sql` ya tienen la estructura.

## Fuentes oficiales (2023)
- Preconteo: https://resultadosprec2023.registraduria.gov.co/
- Escrutinio: https://escrutinios-2023.registraduria.gov.co/
- Formularios E-14: https://divulgacione14.registraduria.gov.co/app/search
- Histórico consolidado: https://www.registraduria.gov.co/-Historico-de-resultados-electorales-.html

Durante la construcción, los pedidos automáticos a los portales de la Registraduría
se bloquearon (HTTP 403 o permiso denegado). Hacen falta descarga manual o una autorización explícita.

## Procedimiento sugerido
1. Descargar los archivos por departamento (escrutinio, no preconteo) y guardarlos en `raw/mesa_2023/`.
2. Mapear los códigos de la Registraduría (departamento 2 dígitos + municipio 3 dígitos, con
   numeración propia) a DIVIPOLA. El servicio Esri trae ambos códigos
   (`codigoMunicipalDivipolCompleto` → `codigoMunicipalMGNCompleto`), así que sirve como tabla de equivalencia.
3. Cargar en `fact_resultados_mesa` con `etapa_resultado = 'Escrutinio'` y el identificador del formulario.
4. No asignar geometrías a puestos sin geocodificación verificada.
5. Con el total municipal del escrutinio, recalcular los válidos, el margen 2023 y el NEP 2023, que hoy no están disponibles.
