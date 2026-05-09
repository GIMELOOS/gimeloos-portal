# 📋 Documentación técnica — Portal GIMELOOS

> Documento para desarrollador externo. Última actualización: mayo 2026.

---

## 1. ¿Qué es este proyecto?

Portal web privado para la empresa **GIMELOOS** (campamentos y experiencias juveniles). Permite a las familias de los participantes:

- Consultar la información de su viaje/campamento
- Subir documentación obligatoria
- Realizar y justificar pagos
- Ver el itinerario y checklist de equipaje
- Enviar preguntas al equipo de GIMELOOS
- Recibir notificaciones en tiempo real (email + in-app)

El panel de administración permite al equipo de GIMELOOS gestionar participantes, documentos, pagos, preguntas, viajes y plantillas.

---

## 2. Stack tecnológico

| Tecnología | Versión | Uso |
|---|---|---|
| **Next.js** | 16.2.0 | Framework full-stack (App Router) |
| **React** | 19.2.4 | UI |
| **Tailwind CSS** | v4 | Estilos |
| **shadcn/ui** | latest | Componentes UI (Button, Card, Badge, etc.) |
| **Framer Motion** | 12.x | Animaciones |
| **Supabase** | 2.x | Base de datos PostgreSQL + Auth |
| **Google Drive API** | v3 (OAuth2) | Almacenamiento de archivos |
| **Resend** | 6.x | Envío de emails transaccionales |
| **XLSX** | 0.18 | Importación de participantes desde Excel |
| **Lucide React** | 0.577 | Iconos |

---

## 3. Estructura de archivos clave

```
gimeloos-portal/
├── app/
│   ├── page.js                          # Punto de entrada → renderiza portal-app.jsx
│   ├── layout.js                        # Layout global
│   ├── globals.css                      # Estilos globales
│   ├── portal-app.jsx                   # ⭐ FICHERO PRINCIPAL (~3.000 líneas)
│   ├── ui/
│   │   └── calculadora-campamento.jsx   # Calculadora de precios integrada
│   └── api/
│       ├── notify/route.js              # Envío de emails (Resend) + notificaciones in-app
│       ├── upload-to-drive/route.js     # Subida de archivos a Google Drive
│       ├── import-excel/route.js        # Importación de participantes desde Excel
│       ├── create-auth-users/route.js   # Creación de usuarios en Supabase Auth
│       ├── invite-participant/route.js  # Invitación a participantes
│       └── auth/
│           ├── google-setup/route.js    # Inicia OAuth2 con Google
│           └── google-callback/route.js # Callback OAuth2, guarda token
├── lib/
│   └── supabase.js                      # Cliente de Supabase
├── components/ui/                       # Componentes shadcn/ui
├── .env.local                           # Variables de entorno (NO subir a git)
├── .google-token.json                   # Token OAuth2 Google (NO subir a git)
└── .claude/
    └── launch.json                      # Config arranque servidor (autoPort: false)
```

---

## 4. Variables de entorno (.env.local)

```env
# Supabase
NEXT_PUBLIC_SUPABASE_URL=https://lggxdgkzhaszcohnrbeo.supabase.co
NEXT_PUBLIC_SUPABASE_ANON_KEY=eyJ...
SUPABASE_SERVICE_ROLE_KEY=eyJ...

# Google Drive (OAuth2 con cuenta personal)
GOOGLE_OAUTH_CLIENT_ID=...
GOOGLE_OAUTH_CLIENT_SECRET=...
GOOGLE_DRIVE_FOLDER_ID=1jAoMMJoE5J042GvEbs904BygtyO7_wnW
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n..."
GOOGLE_SERVICE_ACCOUNT_EMAIL=...

# Resend (emails transaccionales)
RESEND_API_KEY=re_...
NOTIFY_FROM=onboarding@resend.dev   # Cambiar a updates@gimeloos.com cuando el dominio esté verificado
ADMIN_EMAIL=info@gimeloos.com
```

> ⚠️ El archivo `.env.local` y `.google-token.json` **nunca deben subirse a Git**. Están en `.gitignore`.

---

## 5. Cómo arrancar en local

```bash
# 1. Clonar el repositorio
git clone <url-del-repo>
cd gimeloos-portal

# 2. Instalar dependencias
npm install

# 3. Crear .env.local con las variables (ver sección 4)

# 4. Conectar Google Drive (primera vez)
# Iniciar el servidor y visitar:
http://localhost:3000/api/auth/google-setup
# Seguir el flujo OAuth2 → se guarda .google-token.json

# 5. Arrancar servidor de desarrollo
npm run dev
# → http://localhost:3000
```

> ⚠️ El callback de OAuth2 de Google está hardcodeado a `http://localhost:3000`. Por eso el servidor **debe usar el puerto 3000** (`autoPort: false` en `.claude/launch.json`).

---

## 6. Base de datos — Tablas de Supabase

### `participants`
| Columna | Tipo | Descripción |
|---|---|---|
| id | uuid (PK) | ID del participante (= auth.uid()) |
| email | text | Email de acceso |
| username | text | Nombre de usuario para login |
| participant_name | text | Nombre real del participante |
| role | text | `'client'` o `'admin'` |
| trip_id | uuid (FK) | Viaje asignado |
| checklist_state | jsonb | Estado del checklist de equipaje |
| grupo_origen | text | Grupo de origen del participante |

### `trips`
| Columna | Tipo | Descripción |
|---|---|---|
| id | uuid (PK) | |
| name | text | Nombre del viaje/campamento |
| dates | text | Fechas del viaje |
| location | text | Destino |
| description | text | Descripción |
| meeting_point | text | Punto de encuentro |
| itinerary | jsonb | Array de bloques del itinerario |
| checklist | jsonb | Array de items del checklist |
| hero_images | jsonb | URLs de imágenes del hero |

### `participant_documents`
| Columna | Tipo | Descripción |
|---|---|---|
| id | uuid (PK) | |
| participant_id | uuid (FK) | |
| document_template_id | uuid (FK) | |
| status | text | `pending` / `pending_confirmation` / `confirmed` / `rejected` |
| uploaded_file_name | text | Nombre del archivo subido |
| drive_url | text | URL del archivo en Google Drive |
| confirmed_at | timestamptz | Fecha de confirmación |

### `participant_payments`
| Columna | Tipo | Descripción |
|---|---|---|
| id | uuid (PK) | |
| participant_id | uuid (FK) | |
| payment_key | text | `reservation` / `firstInstallment` / `secondInstallment` |
| name | text | Nombre del pago |
| amount | text | Importe |
| status | text | `pending` / `sent` / `confirmed` |
| proof_name | text | Nombre del justificante |
| proof_path | text | URL del justificante en Google Drive |
| due_date | date | Fecha límite de pago |

### `document_templates`
| Columna | Tipo | Descripción |
|---|---|---|
| id | uuid (PK) | |
| name | text | Nombre del documento |
| description | text | Descripción |
| trip_id | uuid (FK) | Viaje al que aplica |
| drive_url | text | URL de la plantilla en Google Drive |

### `participant_questions`
| Columna | Tipo | Descripción |
|---|---|---|
| id | uuid (PK) | |
| participant_id | uuid (FK) | |
| message | text | Pregunta del participante |
| status | text | `sent` / `read` / `replied` |
| reply | text | Respuesta del equipo |
| replied_at | timestamptz | Fecha de respuesta |

### `notifications`
| Columna | Tipo | Descripción |
|---|---|---|
| id | uuid (PK) | |
| participant_id | uuid (FK) | |
| type | text | Tipo de notificación (ver sección 8) |
| title | text | Título visible in-app |
| body | text | Cuerpo visible in-app |
| read | boolean | Si ha sido leída |
| created_at | timestamptz | |

---

## 7. Google Drive — Estructura de carpetas

Los archivos se organizan automáticamente así:

```
📁 Carpeta raíz (GOOGLE_DRIVE_FOLDER_ID)
└── 📁 Nombre del campamento (ej: "Camp Dolomitas 2026")
    └── 📁 Nombre del participante (ej: "Lucas Alexander")
        ├── 📁 documentos
        │   └── dni.pdf, autorizacion.pdf...
        └── 📁 pagos
            └── justificante_reserva.jpg...
```

La autenticación es via **OAuth2** (no service account). El token se guarda en `.google-token.json` y se refresca automáticamente.

---

## 8. Sistema de notificaciones

### Tipos de notificación disponibles

| Tipo | Destinatario | Descripción |
|---|---|---|
| `doc_confirmed` | Participante | Documento confirmado por admin |
| `doc_rejected` | Participante | Documento rechazado por admin |
| `payment_confirmed` | Participante | Pago confirmado por admin |
| `question_replied` | Participante | Admin ha respondido una pregunta |
| `payment_reminder` | Participante | Recordatorio de pago próximo (7, 3, 1 días antes) |
| `doc_reminder` | Participante | Recordatorio de documento pendiente |
| `admin_doc_uploaded` | Admin | Nuevo documento subido por participante |
| `admin_payment_uploaded` | Admin | Nuevo justificante subido por participante |
| `admin_new_question` | Admin | Nueva pregunta de un participante |

### Cómo funciona
- Cada notificación envía un **email via Resend** + guarda una **notificación in-app** en la tabla `notifications`
- El participante ve las notificaciones no leídas con un badge rojo en la campana (navbar)
- Los acordeones de Documentación, Pagos y Preguntas muestran un **punto rojo** si hay notificaciones no leídas de ese tipo
- Los recordatorios de pago se disparan por milestones (7, 3, 1 días) y se guardan en `localStorage` para no repetirse

---

## 9. Componentes principales (portal-app.jsx)

| Componente | Descripción |
|---|---|
| `LoginScreen` | Pantalla de inicio de sesión |
| `ClientPortal` | Portal del participante (documentos, pagos, checklist, preguntas, notificaciones) |
| `ClientDocuments` | Sección de documentos del participante con barra de progreso de subida |
| `ClientPayments` | Sección de pagos con barra de progreso de subida |
| `ClientQuestions` | Historial de preguntas y formulario de nueva pregunta |
| `ClientChecklist` | Checklist interactivo de equipaje |
| `ClientItinerary` | Vista del itinerario del viaje |
| `HeroBanner` | Banner principal con imágenes del campamento |
| `AccordionSection` | Acordeón reutilizable con soporte de badge de no leídos |
| `AdminPanel` | Panel completo de administración |
| `AdminClients` | Gestión de participantes (CRUD, importación Excel, asignación de grupos) |
| `AdminDocs` | Gestión de plantillas de documentos y revisión de documentos subidos |
| `AdminPayments` | Gestión y confirmación de pagos |
| `AdminQuestions` | Vista y respuesta de preguntas de participantes |
| `AdminTrips` | CRUD de viajes (itinerario, checklist, imágenes) |
| `usePaymentReminders` | Hook que gestiona recordatorios automáticos de pago |

---

## 10. Tareas pendientes (por orden de prioridad)

### 🔴 Críticas

1. **Activar Row Level Security (RLS) en Supabase**
   - Ejecutar el SQL comentado al inicio de `portal-app.jsx` en el SQL Editor de Supabase
   - Sin RLS, cualquier usuario autenticado puede leer datos de otros participantes

2. **Verificar dominio en Resend**
   - El dominio `gimeloos.com` está registrado en **Hostalia**
   - Añadir los 4 registros DNS que proporciona Resend (TXT DKIM, MX send, TXT SPF, TXT DMARC)
   - Una vez verificado, cambiar `NOTIFY_FROM` en `.env.local` de `onboarding@resend.dev` a `updates@gimeloos.com`
   - El correo de Google Workspace (`info@gimeloos.com`) no se ve afectado — los registros van en subdominios

3. **Renovar token de Google Drive**
   - El archivo `.google-token.json` contiene el refresh token OAuth2
   - Si expira o se revoca, visitar `http://localhost:3000/api/auth/google-setup` para reconectar

### 🟡 Mejoras pendientes

4. **Filtro "Grupo origen" en panel admin**
   - El select de filtro por grupo de origen no filtra correctamente la lista de participantes

5. **Despliegue en producción**
   - El portal actualmente solo corre en local (`localhost:3000`)
   - Recomendación: desplegar en **Vercel** (gratuito para proyectos Next.js)
   - Al desplegar, actualizar el callback de Google OAuth2 a la URL de producción en Google Cloud Console
   - Añadir todas las variables de `.env.local` como variables de entorno en Vercel

6. **Supabase Auth completo**
   - Actualmente el login compara contra la tabla `participants` directamente
   - Migrar a Supabase Auth nativo (`supabase.auth.signInWithPassword`) para activar RLS automáticamente

7. **Contraseñas seguras**
   - Implementar reset de contraseña por email
   - Asegurarse de que las contraseñas se guardan como hash en Supabase Auth (no en texto plano en la tabla `participants`)

### 🟢 Mejoras deseables

8. **Notificaciones push** (opcional)
   - Actualmente las notificaciones in-app solo se cargan al iniciar sesión
   - Implementar polling o Supabase Realtime para notificaciones en tiempo real

9. **Exportación de datos**
   - Permitir exportar listado de participantes, estado de documentos y pagos a Excel desde el panel admin

---

## 11. Credenciales y accesos (pasar de forma segura al desarrollador)

| Servicio | Dónde está |
|---|---|
| Supabase (URL + keys) | `.env.local` |
| Google Drive OAuth2 | `.env.local` + `.google-token.json` |
| Resend API Key | `.env.local` |
| Google Cloud Console | Cuenta Google del propietario (para gestionar OAuth2) |
| Hostalia (DNS) | Cuenta del propietario |
| Supabase Dashboard | [supabase.com](https://supabase.com) — proyecto `lggxdgkzhaszcohnrbeo` |

---

*Documentación generada en mayo 2026. Para cualquier duda sobre la arquitectura o decisiones técnicas tomadas, revisar los comentarios inline en `portal-app.jsx`.*
