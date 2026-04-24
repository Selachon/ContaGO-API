# ContaGO-API

Backend API de ContaGO (Express + TypeScript).

## Causacion (PDF + Excel + Drive)

Se implemento el endpoint `POST /documents/causation/build-and-upload` para:

1. Recibir PDF inicial + Excel
2. Buscar coincidencia exacta contra columna `X` usando nombre del PDF
3. Tomar link de cuenta de cobro en columna `L`
4. Combinar ambos PDFs (inicial primero, cuenta de cobro segundo)
5. Subir PDF final a Drive en estructura `raiz/año/mes`

Columnas usadas:

- `B`: fecha documento (año/mes destino)
- `L`: enlace de Drive de cuenta de cobro
- `X`: referencia (clave de búsqueda y nombre final)

Regla de comparación:

- sin extensión `.pdf`
- trim de espacios
- case-insensitive

## Causacion desde Google Sheet fijo

Se implemento un flujo alterno (sin adjuntar Excel por request) con fuente fija en Google Sheets:

- Endpoint: `POST /causation/build`
- Endpoint temporal de diagnóstico: `POST /causation/test-openai-file`
- Fuente: Google Sheet `Registro de Cuentas de Cobro`
  - `spreadsheetId`: `1SHCCqtbesErScljl7UqaSeQ2oKBzHRGzYttIU4H70ns`
  - `gid`: `42421166`

Cruce de columnas:

- `B`: fecha documento -> carpeta año/mes destino en Drive
- `L`: enlace de Drive de cuenta de cobro
- `X`: referencia de factura -> clave de cruce con nombre del PDF inicial

Comparación de referencia:

- quitar `.pdf`
- trim de espacios
- case-insensitive

El endpoint `POST /causation/build` acepta:

- `application/json` con `openaiFileIdRefs` (modo recomendado para ChatGPT Actions)
- `multipart/form-data` con `document` (compatibilidad manual)

En `openaiFileIdRefs`, toma el primer archivo del arreglo como PDF inicial.

`POST /causation/test-openai-file` sirve para validar auth + payload de Actions sin ejecutar Drive/Sheets/merge.

Si `openaiFileIdRefs` falta o está vacío, `POST /causation/build` responde `400` con:

```json
{
  "ok": false,
  "code": "missing_input_file",
  "message": "openaiFileIdRefs debe contener al menos un archivo"
}
```

## Integracion Siigo (Fase 1)

La integracion de Siigo esta implementada en:

- `src/services/siigoService.ts`
- `src/routes/siigo.ts`

Todos los endpoints expuestos por ContaGO estan bajo:

- `/integrations/siigo/*`

## Autenticacion para /integrations/siigo/*

Las rutas de Siigo aceptan dos mecanismos de autenticacion:

1. JWT normal de ContaGO (`Bearer <JWT_CONTAGO>`)
2. API key interna fija para GPT (`Bearer <GPT_INTERNAL_API_KEY>`)

Variable nueva:

```env
GPT_INTERNAL_API_KEY=your-fixed-internal-api-key
```

Comportamiento:

- Si el bearer coincide exactamente con `GPT_INTERNAL_API_KEY`, se autoriza como acceso interno.
- Si no coincide, se intenta validar como JWT con el middleware actual.
- Si no valida por ninguno de los dos caminos, responde `401` con formato consistente.

## Variables de entorno

```env
SIIGO_API_BASE_URL=https://api.siigo.com
SIIGO_PARTNER_ID=SentiidoAI
SIIGO_USERNAME=...
SIIGO_ACCESS_KEY=...
GOOGLE_DRIVE_CAUSATION_ROOT_FOLDER_ID=...
GOOGLE_SHEETS_CAUSATION_SPREADSHEET_ID=1SHCCqtbesErScljl7UqaSeQ2oKBzHRGzYttIU4H70ns
GOOGLE_SHEETS_CAUSATION_GID=42421166
GOOGLE_DRIVE_CLIENT_EMAIL=service-account@project-id.iam.gserviceaccount.com
GOOGLE_DRIVE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
GOOGLE_DRIVE_SHARED_DRIVE_ID=...
GOOGLE_DRIVE_USE_SHARED_DRIVE=false
```

Notas:

- `SIIGO_PARTNER_ID` tiene default `SentiidoAI`.
- Nunca se expone `SIIGO_ACCESS_KEY` ni el token completo en respuestas.
- Para `POST /causation/build`, debes compartir el Google Sheet "Registro de Cuentas de Cobro" con el `GOOGLE_DRIVE_CLIENT_EMAIL` de la service account.

## Estrategia de token y autenticacion

- El token de Siigo se cachea en memoria con expiracion (`tokenCache`).
- Si no existe token valido, se autentica contra `POST /auth` en Siigo.
- Si una solicitud de Siigo responde `401`, se limpia cache, se reautentica una vez y se reintenta una sola vez.
- Si vuelve a fallar (`401` o cualquier otro error), se responde error limpio al cliente (`siigo_unauthorized` o `siigo_request_failed`).

## Endpoints expuestos

Todos requieren autenticacion de ContaGO (JWT) o API key interna GPT.

- `GET /integrations/siigo/health`
- `POST /integrations/siigo/auth`
- `GET /integrations/siigo/invoices`
- `GET /integrations/siigo/invoices/:id`
- `GET /integrations/siigo/invoices/:id/pdf`
- `GET /integrations/siigo/invoices/:id/xml`
- `GET /integrations/siigo/purchases`
- `GET /integrations/siigo/purchases/:id`
- `GET /integrations/siigo/purchases/search`
- `GET /integrations/siigo/payment-receipts`
- `GET /integrations/siigo/payment-receipts/:id`
- `GET /integrations/siigo/payment-receipts/search`
- `GET /integrations/siigo/purchase-document-types`
- `GET /integrations/siigo/purchase-document-types/search`
- `POST /documents/causation/build-and-upload`
- `POST /causation/build`
- `POST /causation/test-openai-file`
- `GET /integrations/siigo/customers`
- `GET /integrations/siigo/customers/:id`
- `GET /integrations/siigo/products/:id`
- `GET /integrations/siigo/document-types`

## Query params soportados

### Invoices

- `created_start`
- `created_end`
- `updated_start`
- `updated_end`
- `name`
- `customer_identification`
- `customer_branch_office`
- `document_id`
- `date_start`
- `date_end`
- `page`
- `page_size`

### Customers

- `identification`
- `branch_office`
- `created_start`
- `created_end`
- `updated_start`
- `updated_end`
- `page`
- `page_size`

### Document types

- `type`

### Purchase document types search

- `query`
- `code`
- `name`
- `description`

### Purchases search

- `id`
- `number`
- `name`
- `supplier_identification`
- `provider_invoice_prefix`
- `provider_invoice_number`
- `date_start`
- `date_end`
- `page`
- `page_size`

### Payment receipts search

- `id`
- `number`
- `name`
- `document_id`
- `date_start`
- `date_end`
- `third_party_identification`
- `page`
- `page_size`

### Fuente y supuestos de parametros

Se validaron contra el blueprint oficial de Apiary (`/api-description-document`) para:

- `GET /v1/invoices{?created_start}`
- `GET /v1/customers{?created_start}`

El blueprint lista explicitamente los filtros de fecha e identificacion, y muestra paginacion en enlaces (`page`/`page_size`) aunque no siempre los enumera en la tabla de parametros. Por compatibilidad se admiten ambos.

### Filtros nativos Siigo vs filtros locales ContaGO

- `GET /integrations/siigo/purchases`: usa filtros nativos enviados a Siigo (`created_start`, `created_end`, `updated_start`, `updated_end`, `page`, `page_size`).
- `GET /integrations/siigo/payment-receipts`: usa filtros nativos enviados a Siigo (`created_start`, `created_end`, `updated_start`, `updated_end`, `page`, `page_size`).
- `GET /integrations/siigo/purchases/search`: usa nativos `page`, `page_size`; aplica localmente `id`, `number`, `name`, `supplier_identification`, `provider_invoice_prefix`, `provider_invoice_number`, `date_start`, `date_end`.
- `GET /integrations/siigo/payment-receipts/search`: usa nativos `created_start`, `created_end`, `updated_start`, `updated_end`, `page`, `page_size`; aplica localmente `id`, `number`, `name`, `document_id`, `date_start`, `date_end`, `third_party_identification`.
- `GET /integrations/siigo/purchase-document-types`: consulta Siigo con filtro nativo fijo `type=FC`.
- `GET /integrations/siigo/purchase-document-types/search`: consulta Siigo con `type=FC` y aplica filtros locales case-insensitive (`query`, `code`, `name`, `description`).

## Ejemplos de respuesta por endpoint

### 1) `GET /integrations/siigo/health`

```json
{
  "ok": true,
  "source": "siigo",
  "configured": true,
  "missing": [],
  "tokenCached": false,
  "baseUrl": "https://api.siigo.com",
  "partnerId": "SentiidoAI",
  "authMode": "jwt"
}
```

`authMode` aparece cuando la solicitud llega autenticada (por ejemplo `jwt` o `internal_api_key`).

### 2) `POST /integrations/siigo/auth`

```json
{
  "ok": true,
  "source": "siigo",
  "authenticated": true,
  "expiresAt": "2026-04-10T18:20:31.000Z"
}
```

### 3) `GET /integrations/siigo/invoices`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "results": [],
    "pagination": {
      "page": 1,
      "page_size": 25,
      "total_results": 0
    }
  }
}
```

### 4) `GET /integrations/siigo/invoices/:id`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "id": "b446fbc3-99c5-4c2d-8a28-1c0d3721ab91",
    "name": "FV-1-123"
  }
}
```

### 5) `GET /integrations/siigo/invoices/:id/pdf`

Comportamiento:

- Si Siigo responde binario: se devuelve archivo.
- Si Siigo responde JSON con `base64`: se decodifica y se devuelve archivo.
- Si Siigo responde JSON con URL (`url`, `link`, `href`, `download_url`): se devuelve JSON con `download_url`.

Ejemplo JSON cuando Siigo retorna URL:

```json
{
  "ok": true,
  "source": "siigo",
  "message": "Siigo devolvió una URL de descarga en lugar de archivo binario.",
  "data": {
    "download_url": "https://..."
  }
}
```

### 6) `GET /integrations/siigo/invoices/:id/xml`

Mismo comportamiento que PDF (binario/base64/url).

### 7) `GET /integrations/siigo/purchases/:id`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "id": "purchase-id"
  }
}
```

### 7.1) `GET /integrations/siigo/purchases`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "results": []
  }
}
```

### 7.2) `GET /integrations/siigo/purchases/search`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "results": [],
    "_filtering": {
      "native": ["page", "page_size"],
      "local": ["id", "number", "name", "supplier_identification", "provider_invoice_prefix", "provider_invoice_number", "date_start", "date_end"]
    }
  }
}
```

### 7.3) `GET /integrations/siigo/payment-receipts`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "results": []
  }
}
```

### 7.4) `GET /integrations/siigo/payment-receipts/:id`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "id": "payment-receipt-id"
  }
}
```

### 7.5) `GET /integrations/siigo/payment-receipts/search`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "results": [],
    "_filtering": {
      "native": ["created_start", "created_end", "updated_start", "updated_end", "page", "page_size"],
      "local": ["id", "number", "name", "document_id", "date_start", "date_end", "third_party_identification"]
    }
  }
}
```

### 8) `GET /integrations/siigo/customers`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "results": []
  }
}
```

### 9) `GET /integrations/siigo/customers/:id`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "id": "customer-id"
  }
}
```

### 10) `GET /integrations/siigo/products/:id`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "id": "product-id"
  }
}
```

### 11) `GET /integrations/siigo/document-types`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "results": []
  }
}
```

### 12) `GET /integrations/siigo/purchase-document-types`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "results": [
      {
        "id": 11,
        "code": "DS",
        "name": "Documento Soporte"
      }
    ]
  }
}
```

### 13) `GET /integrations/siigo/purchase-document-types/search`

```json
{
  "ok": true,
  "source": "siigo",
  "data": {
    "results": [
      {
        "id": 11,
        "code": "DS",
        "name": "Documento Soporte",
        "description": "Doc soporte proveedores",
        "possible_match": true
      }
    ],
    "total_results": 1,
    "applied_filters": {
      "query": "soporte",
      "code": null,
      "name": null,
      "description": null
    }
  }
}
```

### 14) `POST /documents/causation/build-and-upload`

```json
{
  "ok": true,
  "data": {
    "reference": "DS-2025-000123",
    "matched_row": 17,
    "drive_source_link": "https://drive.google.com/file/d/1Abc.../view",
    "year_folder": "2025",
    "month_folder": "01-Enero",
    "uploaded_file_name": "DS-2025-000123.pdf",
    "uploaded_file_id": "1Zxy...",
    "uploaded_file_url": "https://drive.google.com/file/d/1Zxy.../view"
  }
}
```

### 15) `POST /causation/build`

```json
{
  "ok": true,
  "data": {
    "reference": "DS-2025-000123",
    "matched_row": 17,
    "registro_source": {
      "spreadsheetId": "1SHCCqtbesErScljl7UqaSeQ2oKBzHRGzYttIU4H70ns",
      "gid": "42421166"
    },
    "drive_source_link": "https://drive.google.com/file/d/1Abc.../view",
    "year_folder": "2025",
    "month_folder": "01-Enero",
    "uploaded_file_name": "DS-2025-000123.pdf",
    "uploaded_file_id": "1Zxy...",
    "uploaded_file_url": "https://drive.google.com/file/d/1Zxy.../view"
  }
}
```

## Formato de errores

Error estandar:

```json
{
  "ok": false,
  "source": "siigo",
  "code": "siigo_request_failed",
  "message": "Mensaje de error",
  "details": {}
}
```

Codigos comunes:

- `siigo_config_missing`
- `siigo_auth_failed`
- `siigo_auth_invalid_response`
- `siigo_network_error`
- `siigo_unauthorized`
- `siigo_not_found`
- `siigo_request_failed`

## Comportamientos esperados

- `POST /auth` nunca devuelve el `access_token` de Siigo.
- `GET /health` nunca devuelve credenciales ni token.
- `GET /health` puede incluir `authMode` (`jwt` o `internal_api_key`) sin exponer secretos.
- Se envia `Partner-Id` en autenticacion y en todas las llamadas hacia Siigo.
- Ante `401` de Siigo, hay un solo reintento con reautenticacion.
- En `pdf/xml`, la API soporta respuestas de Siigo en binario o JSON (`base64` o URL).
- Si Siigo responde error, la API responde con formato de error consistente.

## Limitaciones conocidas

- Cache de token en memoria de proceso: en despliegues con multiples replicas, cada replica maneja su propio token.
- No hay circuit breaker ni backoff automatico para `429`/`503` en esta fase.
- Siigo puede cambiar estructuras de respuesta en `pdf/xml`; se maneja compatibilidad basica, pero pueden requerirse ajustes puntuales.
- No se versiona aun el contrato de respuestas de ContaGO para este modulo (se recomienda en fase 2).
- El flujo de causacion requiere Google Drive conectado para el usuario y acceso al archivo enlazado en columna L.
- El flujo de causacion no usa OCR ni lectura de contenido PDF; solo cruza por nombre de archivo vs columna X.
- El endpoint `POST /causation/build` depende de permisos de la service account sobre el Sheet y sobre los archivos/carpeta en Drive.
- Si `GOOGLE_DRIVE_USE_SHARED_DRIVE=true`, se debe configurar `GOOGLE_DRIVE_SHARED_DRIVE_ID`.

## Ejemplos curl actualizados

```bash
# Login ContaGO para obtener JWT
curl -X POST http://localhost:8000/auth/login \
  -H "Content-Type: application/json" \
  -d '{"email":"admin@contago.com","password":"tu-clave"}'
```

```bash
# Health Siigo
curl http://localhost:8000/integrations/siigo/health \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Health Siigo con API key interna (GPT)
curl http://localhost:8000/integrations/siigo/health \
  -H "Authorization: Bearer <GPT_INTERNAL_API_KEY>"
```

```bash
# Autenticar cache de Siigo
curl -X POST http://localhost:8000/integrations/siigo/auth \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Listar facturas
curl "http://localhost:8000/integrations/siigo/invoices?created_start=2025-01-01&created_end=2025-01-31&page=1&page_size=25" \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Detalle de factura
curl http://localhost:8000/integrations/siigo/invoices/<INVOICE_ID> \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# PDF de factura
curl -L http://localhost:8000/integrations/siigo/invoices/<INVOICE_ID>/pdf \
  -H "Authorization: Bearer <JWT_CONTAGO>" \
  -o factura.pdf
```

```bash
# XML de factura
curl -L http://localhost:8000/integrations/siigo/invoices/<INVOICE_ID>/xml \
  -H "Authorization: Bearer <JWT_CONTAGO>" \
  -o factura.xml
```

```bash
# Listar clientes
curl "http://localhost:8000/integrations/siigo/customers?identification=900123456&page=1&page_size=25" \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Tipos de comprobante
curl "http://localhost:8000/integrations/siigo/document-types?type=FV" \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Listar facturas de compra
curl "http://localhost:8000/integrations/siigo/purchases?page=1&page_size=25" \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Buscar facturas de compra (filtros locales + paginacion nativa)
curl "http://localhost:8000/integrations/siigo/purchases/search?supplier_identification=900123456&provider_invoice_prefix=PF&provider_invoice_number=7788&page=1&page_size=25" \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Listar recibos de pago/egreso
curl "http://localhost:8000/integrations/siigo/payment-receipts?created_start=2025-01-01&page=1&page_size=25" \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Buscar recibos de pago/egreso
curl "http://localhost:8000/integrations/siigo/payment-receipts/search?document_id=321&third_party_identification=800123&page=1&page_size=25" \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Tipos de comprobante de compras (type=FC en Siigo)
curl "http://localhost:8000/integrations/siigo/purchase-document-types" \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Buscar posible Documento Soporte (DS)
curl "http://localhost:8000/integrations/siigo/purchase-document-types/search?query=soporte" \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
# Construir y subir archivo de causacion
curl -X POST "http://localhost:8000/documents/causation/build-and-upload" \
  -H "Authorization: Bearer <JWT_CONTAGO>" \
  -F "initial_pdf=@/ruta/DS-2025-000123.pdf" \
  -F "excel=@/ruta/Cruce-Causacion.xlsx" \
  -F "debug=true"
```

```bash
# Construir causacion usando Google Sheet fijo (Registro de Cuentas de Cobro)
curl -X POST "http://localhost:8000/causation/build" \
  -H "Authorization: Bearer <JWT_CONTAGO>" \
  -F "document=@/ruta/DS-2025-000123.pdf" \
  -F "debug=true"
```

```bash
# Construir causacion con ChatGPT Actions (openaiFileIdRefs)
curl -X POST "http://localhost:8000/causation/build" \
  -H "Authorization: Bearer <GPT_INTERNAL_API_KEY>" \
  -H "Content-Type: application/json" \
  -d '{
    "openaiFileIdRefs": [
      {
        "name": "DS-1-1570.pdf",
        "id": "file-abc123",
        "mime_type": "application/pdf",
        "download_link": "https://files.openai.com/v1/files/file-abc123/content"
      }
    ],
    "debug": true
  }'
```

```bash
# Validar payload openaiFileIdRefs sin ejecutar flujo completo
curl -X POST "http://localhost:8000/causation/test-openai-file" \
  -H "Authorization: Bearer <GPT_INTERNAL_API_KEY>" \
  -H "Content-Type: application/json" \
  -d '{
    "openaiFileIdRefs": [
      {
        "name": "DS-1-1570.pdf",
        "id": "file-abc123",
        "mime_type": null,
        "download_link": "https://files.openai.com/v1/files/file-abc123/content"
      }
    ]
  }'
```

## Pruebas automaticas minimas

Se agregaron pruebas para:

- `health`
- `auth`
- `invoices list`
- `purchases list`
- `purchases search`
- `payment-receipts search`
- `purchase-document-types`
- `purchase-document-types search`
- `causation match/date/filename`
- `causation rows matching` (exacta, sin coincidencia, múltiples, columna L vacía)
- `causation date folders` (año/mes)
- `causation openaiFileIdRefs` (válido, faltante, vacío, no PDF)
- `causation test-openai-file` (auth + payload mínimo)
- `retry tras 401`
- `error por configuracion faltante`

Ejecutar:

```bash
npm run test:siigo
npm run test:causation
```

## Prueba local y Railway

### Local

1. Completar variables `SIIGO_*` en `.env`.
2. (Opcional GPT) Configurar `GPT_INTERNAL_API_KEY`.
3. Ejecutar `npm run dev`.
4. Obtener JWT en `/auth/login` o usar API key interna.
5. Probar endpoints `/integrations/siigo/*`.

### Railway

1. Crear servicio en Railway apuntando a este backend.
2. Definir `SIIGO_*` en Variables del servicio.
3. Definir `GPT_INTERNAL_API_KEY` en Railway para consumo GPT.
4. Desplegar.
5. Probar:

```bash
curl https://<tu-dominio-railway>/integrations/siigo/health \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```

```bash
curl https://<tu-dominio-railway>/integrations/siigo/health \
  -H "Authorization: Bearer <GPT_INTERNAL_API_KEY>"
```

## Configuracion sugerida para GPT Action

1. En tu GPT Action, define `Server URL` hacia tu API (por ejemplo `https://<tu-dominio-railway>`).
2. En autenticacion, usa `API Key` con header `Authorization`.
3. Valor del secreto: `Bearer <GPT_INTERNAL_API_KEY>`.
4. Limita el scope de acciones solo a endpoints necesarios en `/integrations/siigo/*`.
5. Rota periodicamente `GPT_INTERNAL_API_KEY` desde variables de entorno de Railway.
