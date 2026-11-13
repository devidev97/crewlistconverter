# ADR Crewlist

Web local para convertir listas de tripulación en `.xls`, `.xlsx` o `.csv` al formato **Crew CSV** de 29 columnas usado por la Autoridad Portuaria.

## Uso con Vite

Instala las dependencias e inicia el servidor de desarrollo:

```bash
npm install
npm run dev
```

Vite escuchará en `0.0.0.0:5173`, por lo que también funcionará detrás del proxy de code-server. El dominio `code-server.devidev.es` y sus subdominios están permitidos en `vite.config.mjs`.

Si Vite ya estaba en ejecución antes de cambiar la configuración, detenlo con `Ctrl+C` y vuelve a iniciarlo con `npm run dev`.

El acceso es el mismo que en ADR: autenticación básica con `BASIC_AUTH_USER` y `BASIC_AUTH_PASSWORD`. Copia `.env.example` a `.env` y pon los mismos valores. Sin esas variables la aplicación responde que no está disponible. En Vercel hay que definir las dos variables de entorno.

Después:

1. Carga el Excel recibido.
2. Revisa los tripulantes marcados con incidencias.
3. Completa la ruta y las horas.
4. Añade o duplica movimientos si necesitas varias salidas.
5. Descarga un CSV o, si hay varios movimientos, el paquete ZIP.

Para comprobar la versión de producción:

```bash
npm run build
npm run preview
```

## Formato generado

- Separador: punto y coma (`;`).
- 29 columnas, en el mismo orden que los CSV de referencia.
- Fechas personales: `DD/MM/AAAA`.
- Fechas de movimiento: `DD/MM/AAAA HH:MM`.
- Final de línea: Windows CRLF.
- Codificación: UTF-8 sin BOM.
- Nombre: `Lista de tripulantes.csv`.

La descripción de categoría se normaliza según su código (`ZOT` → `Otros`, `ZCP` → `Capitan`, etc.). Las columnas no presentes en el Excel, como `SchengenArea`, se incluyen vacías.

## Privacidad

Toda la lectura, validación y conversión se realiza en el navegador. Los archivos y datos personales no se suben a ningún servidor.

## Prueba de conversión

Con Node.js disponible:

```bash
node tests/converter.test.js
```

La prueba genera manifiestos sintéticos para comprobar las 29 columnas, los puertos, las fechas y el nombre de salida sin almacenar datos personales en el repositorio.

## Dependencia incluida

`public/vendor/xlsx.full.min.js` es SheetJS Community Edition 0.18.5. Su licencia Apache 2.0 está en `public/vendor/LICENSE`.
