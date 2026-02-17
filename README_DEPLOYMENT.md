# 🚀 Guía de Despliegue - Dashboard Equipos I&C

Esta guía te ayudará a desplegar tu aplicación en Streamlit Community Cloud para que sea accesible públicamente mediante una URL.

---

## 📋 Requisitos Previos

1. ✅ Cuenta de GitHub (gratis)
2. ✅ Cuenta de Streamlit Community Cloud (gratis)
3. ✅ Archivos del proyecto listos

---

## 🗂️ Archivos Necesarios (Ya Creados)

```
dashboard-equipos-ic/
├── app_dashboard_ic.py          ← Aplicación principal
├── requirements.txt              ← Dependencias (CREADO)
├── .gitignore                    ← Archivos a ignorar (CREADO)
├── .streamlit/
│   └── config.toml              ← Configuración (CREADO)
├── README.md                     ← Documentación
└── README_DEPLOYMENT.md         ← Esta guía
```

---

## 📝 PASO 1: Crear Cuenta en GitHub

### Si NO tienes cuenta:
1. Ve a: https://github.com
2. Haz clic en **"Sign up"**
3. Ingresa tu email
4. Crea una contraseña
5. Elige un nombre de usuario
6. Verifica tu cuenta por email

### Si YA tienes cuenta:
- Inicia sesión en: https://github.com

---

## 📦 PASO 2: Crear Repositorio en GitHub

### Opción A: Desde la Web (Recomendado)

1. **Ir a GitHub:**
   - https://github.com
   - Haz clic en el botón **"+"** (arriba derecha)
   - Selecciona **"New repository"**

2. **Configurar repositorio:**
   ```
   Repository name: dashboard-equipos-ic
   Description: Dashboard de seguimiento para proyectos de Instrumentación y Control
   Visibility: Public ✓ (para Streamlit Community Cloud gratuito)

   ☐ Add a README file (no marcar, ya lo tenemos)
   ☐ Add .gitignore (no marcar, ya lo tenemos)
   ☐ Choose a license (opcional)
   ```

3. **Crear repositorio:**
   - Haz clic en **"Create repository"**
   - ⚠️ NO cierres esta página, la necesitarás en el siguiente paso

### Opción B: Desde GitHub Desktop

1. Descarga GitHub Desktop: https://desktop.github.com
2. Instala y abre la aplicación
3. File → New Repository
4. Configura igual que arriba
5. Crea el repositorio

---

## 📤 PASO 3: Subir Archivos a GitHub

### Método 1: Usando la Web (Más Fácil)

1. **En la página de tu repositorio vacío:**
   - Verás: "Quick setup — if you've done this kind of thing before"
   - Haz clic en **"uploading an existing file"**

2. **Arrastrar archivos:**
   - Arrastra TODOS estos archivos a la ventana:
     ```
     ✓ app_dashboard_ic.py
     ✓ requirements.txt
     ✓ .gitignore
     ✓ README.md
     ✓ NOVEDADES_v1.0.4.txt
     ```

3. **Subir carpeta .streamlit:**
   - Necesitas subir config.toml en carpeta .streamlit
   - GitHub Desktop o Git CLI son mejores para esto
   - O crea la carpeta manualmente en GitHub:
     - Click "Add file" → "Create new file"
     - Escribe: `.streamlit/config.toml`
     - Pega el contenido del archivo

4. **Commit:**
   - En "Commit changes"
   - Mensaje: "Initial commit - Dashboard v1.0.4"
   - Haz clic en **"Commit changes"**

### Método 2: Usando Git CLI

Si tienes Git instalado:

```bash
# 1. Inicializar repositorio local
git init
git add .
git commit -m "Initial commit - Dashboard v1.0.4"

# 2. Conectar con GitHub (reemplaza con tu URL)
git remote add origin https://github.com/TU_USUARIO/dashboard-equipos-ic.git

# 3. Subir archivos
git branch -M main
git push -u origin main
```

### Método 3: Usando GitHub Desktop

1. Abre GitHub Desktop
2. File → Add Local Repository
3. Selecciona la carpeta de tu proyecto
4. Commit to main: "Initial commit"
5. Publish repository

---

## ☁️ PASO 4: Crear Cuenta en Streamlit Community Cloud

1. **Ve a:**
   - https://streamlit.io/cloud

2. **Regístrate:**
   - Haz clic en **"Sign up"**
   - **Opción recomendada:** "Continue with GitHub"
   - Autoriza a Streamlit a acceder a tu cuenta de GitHub

3. **Completa perfil:**
   - Nombre
   - Email (se completa automáticamente)

---

## 🚀 PASO 5: Desplegar la Aplicación

1. **En Streamlit Cloud Dashboard:**
   - Haz clic en **"New app"** o **"Create app"**

2. **Configurar deployment:**
   ```
   Repository: TU_USUARIO/dashboard-equipos-ic
   Branch: main
   Main file path: app_dashboard_ic.py

   App URL (Custom subdomain): equipos-ic-dashboard
   (Opcional - o deja el generado automáticamente)
   ```

3. **Advanced settings (Opcional):**
   - Python version: 3.11 (recomendado)
   - Secrets: No necesario por ahora

4. **Deploy:**
   - Haz clic en **"Deploy!"**
   - Espera 2-5 minutos mientras se despliega

5. **¡Listo!**
   - Tu app estará disponible en:
   - `https://equipos-ic-dashboard.streamlit.app`
   - o
   - `https://TU_USUARIO-dashboard-equipos-ic.streamlit.app`

---

## 🔗 PASO 6: Obtener y Compartir URL

Una vez desplegada, verás:

```
🎉 Your app is live at:
https://equipos-ic-dashboard.streamlit.app
```

**Puedes compartir esta URL con:**
- Tu equipo
- Clientes
- Gerencia
- Contratistas
- Cualquier persona con internet

**La app es pública, cualquiera con la URL puede acceder.**

---

## 🔄 Actualizar la Aplicación

Cuando hagas cambios al código:

### Método 1: GitHub Web
1. Ve a tu repositorio en GitHub
2. Click en `app_dashboard_ic.py`
3. Click en el ícono de lápiz (editar)
4. Haz los cambios
5. Commit changes
6. **Streamlit Cloud detecta el cambio automáticamente**
7. Tu app se actualizará en 1-2 minutos

### Método 2: Git CLI
```bash
# Haz cambios en tu código local
git add .
git commit -m "Descripción del cambio"
git push

# Streamlit Cloud se actualiza automáticamente
```

### Método 3: GitHub Desktop
1. Haz cambios en tu código
2. Commit to main
3. Push origin
4. Actualización automática

---

## 🔒 Seguridad y Privacidad

### ⚠️ IMPORTANTE: Datos Sensibles

**Tu URL de Google Drive está hardcodeada en el código:**
```python
URL_DEFECTO = "https://drive.google.com/uc?export=download&id=..."
```

**Recomendaciones:**

1. **Asegurar archivo de Google Drive:**
   - El archivo debe tener permisos: "Cualquiera con el enlace puede ver"
   - Solo lectura, no edición

2. **Si el archivo contiene datos sensibles:**
   - Considera usar Streamlit Secrets:
   ```python
   # En el código:
   URL_DEFECTO = st.secrets["google_drive_url"]

   # En Streamlit Cloud → Settings → Secrets:
   google_drive_url = "https://drive.google.com/..."
   ```

3. **Alternativa - Autenticación:**
   - Agregar autenticación básica con contraseña
   - Limitar acceso solo a usuarios autorizados

---

## 📊 Monitoreo y Estadísticas

En Streamlit Cloud Dashboard puedes ver:
- **Viewers:** Quién está viendo tu app
- **Activity logs:** Logs de la aplicación
- **Resource usage:** Uso de CPU/memoria
- **Analytics:** Estadísticas de uso

---

## ❓ Solución de Problemas

### Error: "ModuleNotFoundError"
**Causa:** Falta una dependencia en requirements.txt
**Solución:** Agrega la librería faltante y haz push

### Error: "App is not loading"
**Causa:** Error en el código
**Solución:** 
1. Ve a Streamlit Cloud → Tu app → Logs
2. Revisa el error
3. Corrige el código
4. Haz push

### Error: "File not found: Equipos I&C"
**Causa:** El archivo de Google Drive no es accesible
**Solución:**
1. Verifica permisos del archivo
2. Confirma que la URL es correcta
3. Prueba la URL directamente en el navegador

### App muy lenta
**Causa:** Muchos datos o usuarios
**Solución:**
- Optimiza el código
- Reduce cache TTL
- Considera Streamlit Cloud Plus (pago)

---

## 💰 Costos

**Streamlit Community Cloud (Actual):**
- ✅ Gratis
- 1 app privada
- Apps públicas ilimitadas
- Recursos limitados compartidos

**Streamlit Cloud Plus (Opcional):**
- $20-30/mes
- Más recursos
- Apps privadas ilimitadas
- Mejor rendimiento

**Para tu caso:** Community Cloud (gratis) es suficiente.

---

## 📱 Acceso Móvil

La app funciona en:
- ✅ Computadoras (Windows, Mac, Linux)
- ✅ Tablets
- ✅ Smartphones
- ✅ Cualquier dispositivo con navegador

**Responsive:** Streamlit adapta la interfaz automáticamente.

---

## 🎯 Checklist Final

Antes de compartir la URL, verifica:

- [ ] App carga correctamente
- [ ] URL de Google Drive funciona
- [ ] Datos se muestran correctamente
- [ ] Todos los gráficos se ven bien
- [ ] Exportaciones funcionan
- [ ] Filtros funcionan
- [ ] Versión mostrada: 1.0.4

---

## 📞 Soporte

**Problemas con Streamlit Cloud:**
- Documentación: https://docs.streamlit.io/streamlit-community-cloud
- Foro: https://discuss.streamlit.io
- GitHub Issues: https://github.com/streamlit/streamlit/issues

**Problemas con tu app:**
- Revisa los logs en Streamlit Cloud
- Prueba localmente primero: `streamlit run app_dashboard_ic.py`

---

## 🎉 ¡Listo!

Tu dashboard está ahora públicamente accesible. Comparte la URL con tu equipo y comienza a monitorear el progreso de tus equipos I&C desde cualquier lugar.

**URL ejemplo:** 
```
https://equipos-ic-dashboard.streamlit.app
```

**¡Éxito con tu proyecto!** 🏭📊✨
