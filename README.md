# ContaGO-API

Backend API de ContaGO (Express + TypeScript).

## Integracion Siigo (Fase 1)

La integracion de Siigo esta implementada en:

- `src/services/siigoService.ts`
- `src/routes/siigo.ts`

Todos los endpoints expuestos por ContaGO estan bajo:

- `/integrations/siigo/*`

## Variables de entorno

```env
SIIGO_API_BASE_URL=https://api.siigo.com
SIIGO_PARTNER_ID=SentiidoAI
SIIGO_USERNAME=...
SIIGO_ACCESS_KEY=...
```

Notas:

- `SIIGO_PARTNER_ID` tiene default `SentiidoAI`.
- Nunca se expone `SIIGO_ACCESS_KEY` ni el token completo en respuestas.

## Estrategia de token y autenticacion

- El token de Siigo se cachea en memoria con expiracion (`tokenCache`).
- Si no existe token valido, se autentica contra `POST /auth` en Siigo.
- Si una solicitud de Siigo responde `401`, se limpia cache, se reautentica una vez y se reintenta una sola vez.
- Si vuelve a fallar (`401` o cualquier otro error), se responde error limpio al cliente (`siigo_unauthorized` o `siigo_request_failed`).

## Endpoints expuestos

Todos requieren JWT de ContaGO (`Authorization: Bearer <JWT_CONTAGO>`).

- `GET /integrations/siigo/health`
- `POST /integrations/siigo/auth`
- `GET /integrations/siigo/invoices`
- `GET /integrations/siigo/invoices/:id`
- `GET /integrations/siigo/invoices/:id/pdf`
- `GET /integrations/siigo/invoices/:id/xml`
- `GET /integrations/siigo/purchases/:id`
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

### Fuente y supuestos de parametros

Se validaron contra el blueprint oficial de Apiary (`/api-description-document`) para:

- `GET /v1/invoices{?created_start}`
- `GET /v1/customers{?created_start}`

El blueprint lista explicitamente los filtros de fecha e identificacion, y muestra paginacion en enlaces (`page`/`page_size`) aunque no siempre los enumera en la tabla de parametros. Por compatibilidad se admiten ambos.

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
  "partnerId": "SentiidoAI"
}
```

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
- Se envia `Partner-Id` en autenticacion y en todas las llamadas hacia Siigo.
- Ante `401` de Siigo, hay un solo reintento con reautenticacion.
- En `pdf/xml`, la API soporta respuestas de Siigo en binario o JSON (`base64` o URL).
- Si Siigo responde error, la API responde con formato de error consistente.

## Limitaciones conocidas

- Cache de token en memoria de proceso: en despliegues con multiples replicas, cada replica maneja su propio token.
- No hay circuit breaker ni backoff automatico para `429`/`503` en esta fase.
- Siigo puede cambiar estructuras de respuesta en `pdf/xml`; se maneja compatibilidad basica, pero pueden requerirse ajustes puntuales.
- No se versiona aun el contrato de respuestas de ContaGO para este modulo (se recomienda en fase 2).

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

## Pruebas automaticas minimas

Se agregaron pruebas para:

- `health`
- `auth`
- `invoices list`
- `retry tras 401`
- `error por configuracion faltante`

Ejecutar:

```bash
npm run test:siigo
```

## Prueba local y Render

### Local

1. Completar variables `SIIGO_*` en `.env`.
2. Ejecutar `npm run dev`.
3. Obtener JWT en `/auth/login`.
4. Probar endpoints `/integrations/siigo/*`.

### Render

1. Definir `SIIGO_*` en el dashboard de Render.
2. Desplegar.
3. Probar:

```bash
curl https://contago-api.onrender.com/integrations/siigo/health \
  -H "Authorization: Bearer <JWT_CONTAGO>"
```
