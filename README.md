# tsj-kpi-dashboard

Dashboard operativo de TSJ Transporte San José.

## IPK por línea

La solapa **IPK por línea** consulta la sección `ipk` del mismo servicio de datos y espera un objeto como este:

```json
{
  "ok": true,
  "generatedAt": "2026-10-22T12:00:00-03:00",
  "sources": { "ipk": { "name": "Outlook · Reporte IPK" } },
  "ipk": {
    "reportedDate": "21/10/2026",
    "globalDaily": 1.85,
    "globalAccumulated": 1.72,
    "passengers": 125000,
    "productiveKm": 67500,
    "lines": [
      { "linea": "Línea 1", "pasajeros": 12000, "kmProductivos": 6100, "ipkDiario": 1.97, "ipkAcumulado": 1.80 }
    ]
  }
}
```

La fecha de negocio es siempre `reportedDate` (el **Día informado** dentro del PDF). El proceso de ingesta debe ignorar la fecha de recepción del email; en fines de semana y feriados puede recibirse el siguiente día hábil. Si hay dos adjuntos para el mismo `reportedDate`, debe conservarse el del mensaje más reciente.

La fuente es Outlook/Microsoft 365. La automatización debe leer el buzón mediante Microsoft Graph con permiso delegado `Mail.Read` (o permiso de aplicación `Mail.Read` con consentimiento de administrador para un proceso sin usuario), descargar los adjuntos PDF, extraer los campos anteriores y publicar el JSON en la sección `ipk`. No sustituir Outlook por Google Drive.
