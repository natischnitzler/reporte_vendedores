# 🤖 AUTOMATIZACIÓN COMPLETA - Paso a Paso

## ⏱️ Tiempo total: 20 minutos

---

## 📋 PASO 1: Crear repositorio en GitHub (3 min)

### 1.1 Ir a GitHub
- Abre https://github.com/new
- Si no tienes cuenta, créala primero (gratis)

### 1.2 Configurar repositorio
```
Repository name: temponovo-reportes
Description: Reportes automáticos semanales
Visibilidad: ✅ Private (IMPORTANTE: debe ser privado)
Initialize: ✅ Add a README file
```

### 1.3 Crear
- Click en **"Create repository"**
- Copia la URL del repo (ejemplo: `https://github.com/natalia-temp/temponovo-reportes`)

---

## 📁 PASO 2: Subir los archivos al repositorio (5 min)

### 2.1 Descargar archivos necesarios
Descarga estos 4 archivos de Claude:
1. `reporte_vendedores.py` (script principal)
2. `reporte_vendedores.yml` (workflow)
3. `requirements.txt` (dependencias)
4. `AUTOMATIZACION.md` (esta guía)

### 2.2 Subir mediante la web de GitHub

**Opción A - Si sabes usar Git:**
```bash
git clone https://github.com/TU_USUARIO/temponovo-reportes.git
cd temponovo-reportes

# Crear estructura de carpetas
mkdir -p .github/workflows

# Copiar archivos a las carpetas correctas:
# - .github/workflows/reporte_vendedores.yml
# - reporte_vendedores.py (raíz)
# - requirements.txt (raíz)
# - AUTOMATIZACION.md (raíz)

git add .
git commit -m "Setup automatización reportes"
git push
```

**Opción B - Sin Git (más fácil):**
1. Ve a tu repositorio en GitHub
2. Click en **"Add file" → "Upload files"**
3. Arrastra `reporte_vendedores.py`, `requirements.txt`, `AUTOMATIZACION.md`
4. Click **"Commit changes"**
5. Ahora crea la carpeta del workflow:
   - Click **"Add file" → "Create new file"**
   - Nombre: `.github/workflows/reporte_vendedores.yml`
   - Pega el contenido del archivo `reporte_vendedores.yml`
   - Click **"Commit changes"**

---

## 🔐 PASO 3: Configurar Secrets (Contraseñas) (5 min)

### 3.1 Ir a Settings
En tu repositorio: **Settings → Secrets and variables → Actions**

### 3.2 Agregar cada secret
Click en **"New repository secret"** y agregar uno por uno:

| Name | Value | Dónde conseguirlo |
|------|-------|-------------------|
| `ODOO3_URL` | `https://odoo.temponovo.cl` | URL de tu Odoo |
| `ODOO3_DB` | `temponovo` | Nombre base de datos |
| `ODOO3_USER` | `admin` | Usuario Odoo |
| `ODOO3_PASS` | `[tu-contraseña]` | Contraseña Odoo |
| `SMTP_HOST` | `smtp.gmail.com` | Para Gmail |
| `SMTP_PORT` | `587` | Puerto Gmail |
| `SMTP_USER` | `natalia@temponovo.cl` | Email que envía |
| `SMTP_PASS` | `[contraseña-app]`* | Ver abajo ⬇️ |

### 3.3 ⚠️ Contraseña de aplicación Gmail (SMTP_PASS)

**NO uses tu contraseña normal de Gmail.** Necesitas una "contraseña de aplicación":

1. Ve a https://myaccount.google.com/apppasswords
2. Si pide verificación en 2 pasos, actívala primero
3. En "Seleccionar app" → Elige "Correo"
4. En "Seleccionar dispositivo" → Elige "Otro" → Escribe "Reportes Odoo"
5. Click "Generar"
6. Copia la contraseña de 16 caracteres (ej: `abcd efgh ijkl mnop`)
7. Usa esa contraseña en `SMTP_PASS` (sin espacios: `abcdefghijklmnop`)

---

## ✅ PASO 4: Activar el workflow (2 min)

### 4.1 Ir a Actions
En tu repositorio: Click en la pestaña **"Actions"**

### 4.2 Verificar
Deberías ver:
- ✅ "I understand my workflows, go ahead and enable them"
- Click en ese botón verde

Luego verás:
- Workflow: **"Reporte Semanal Vendedores"**

---

## 🧪 PASO 5: Hacer prueba manual (5 min)

### 5.1 Ejecutar manualmente
1. En **Actions** → Click en **"Reporte Semanal Vendedores"**
2. Click en **"Run workflow"** (botón azul a la derecha)
3. Confirma en **"Run workflow"** (verde)

### 5.2 Ver ejecución
- Aparecerá una fila nueva con un círculo amarillo girando 🟡
- Click en esa fila
- Verás los pasos ejecutándose en tiempo real
- Espera 2-4 minutos

### 5.3 Verificar resultados

**Si sale todo bien (✅ verde):**
- Revisa tu email `natalia@temponovo.cl`
- Deberían llegar los reportes de prueba

**Si falla (❌ rojo):**
- Click en el paso que falló
- Lee el error en los logs
- Problemas comunes:
  - Secrets mal configurados → Verifica paso 3
  - Gmail bloqueó el login → Usa contraseña de aplicación
  - Odoo no responde → Verifica URL y credenciales

---

## 📅 PASO 6: Configuración final (1 min)

### 6.1 Cambiar modo producción
Edita `reporte_vendedores.py` en GitHub:
1. Busca la línea: `TEST_MODE = False`
2. Verifica que diga `False` (no `True`)
3. Commit

### 6.2 Verificar horario
En `.github/workflows/reporte_vendedores.yml`:
```yaml
schedule:
  - cron: '0 15 * * 1,3'  # Lunes y Miércoles 11 AM Chile
```

**✅ Listo!** El reporte se enviará automáticamente:
- 🗓️ **Lunes** 11:00 AM
- 🗓️ **Miércoles** 11:00 AM

---

## 🎯 Verificación Final

### ¿Cómo sé que funciona?

1. **Inmediato** (después de la prueba manual):
   - ✅ Llegaron emails a natalia@temponovo.cl
   - ✅ El workflow aparece verde en Actions

2. **Primera ejecución automática** (próximo lunes/miércoles):
   - Verifica que lleguen los emails
   - Revisa en Actions que se ejecutó

### ¿Dónde ver el historial?
- **Actions** → Cada ejecución queda registrada
- Click en cualquier ejecución para ver los logs completos

---

## 🔧 Modificaciones futuras

### Cambiar horario
Edita `.github/workflows/reporte_vendedores.yml`:
```yaml
schedule:
  - cron: 'MINUTO HORA * * DÍA'
```

Ejemplos:
```yaml
'0 13 * * 1,3'     # Lunes y Miércoles 9 AM Chile
'0 19 * * 5'       # Solo Viernes 3 PM Chile  
'0 14 * * 1-5'     # Lunes a Viernes 10 AM Chile
```

🌍 **Importante:** GitHub usa UTC, Chile es UTC-3 o UTC-4:
- Verano (Oct-Mar): 11 AM Chile = 14:00 UTC
- Invierno (Abr-Sep): 11 AM Chile = 15:00 UTC
- Usa `15:00` para cubrir ambos

### Cambiar destinatarios
Edita `reporte_vendedores.py`:
- `VENDEDORES` → lista de vendedores
- `CC_FIJOS` → CC en emails individuales
- `RESUMEN_EMAILS` → destinatarios del resumen ejecutivo

### Pausar temporalmente
**Actions** → **Reporte Semanal Vendedores** → **⋯** → **Disable workflow**

---

## 🆘 Troubleshooting

### ❌ "Authentication failed" en Odoo
**Causa:** Credenciales incorrectas
**Solución:** 
1. Verifica `ODOO3_USER` y `ODOO3_PASS` en Secrets
2. Prueba conectarte manualmente a Odoo con esas credenciales

### ❌ "SMTP authentication error"
**Causa:** Contraseña de Gmail incorrecta
**Solución:**
1. Usa **contraseña de aplicación**, no tu contraseña normal
2. Genera una nueva en https://myaccount.google.com/apppasswords
3. Actualiza `SMTP_PASS` en Secrets

### ❌ No llegan los emails
**Causa:** Revisa spam o configuración SMTP
**Solución:**
1. Busca en spam/promociones
2. Verifica que `SMTP_USER` sea el email correcto
3. En los logs de Actions, busca mensajes de error de SMTP

### ❌ "Workflow not found"
**Causa:** Archivo YML en lugar incorrecto
**Solución:**
- El archivo DEBE estar en `.github/workflows/reporte_vendedores.yml`
- Nota el punto inicial: `.github` (con punto)

### 🕐 Se ejecutó a la hora incorrecta
**Causa:** Diferencia horaria UTC vs Chile
**Solución:**
- Ajusta la hora en el cron
- 11 AM Chile = `15:00` UTC (recomendado para todo el año)

---

## 📞 Soporte

**¿Dudas o problemas?**
- Email: natalia@temponovo.cl
- Logs completos: Actions → [ejecución] → Ver pasos

**Documentación GitHub Actions:**
- https://docs.github.com/en/actions

---

## ✨ Resumen

✅ **Gratis** - 2000 minutos/mes (usa ~5 min por ejecución)  
✅ **Confiable** - Infraestructura de GitHub  
✅ **Sin mantenimiento** - Corre solo  
✅ **Logs completos** - Ves exactamente qué pasó  
✅ **Notificaciones** - Te avisa si falla  

**Próxima ejecución:** Lunes o Miércoles a las 11:00 AM 🎯
