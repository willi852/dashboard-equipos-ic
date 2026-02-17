# 📊 Dashboard de Seguimiento - Equipos I&C

Sistema completo de seguimiento y análisis de avance para proyectos de Instrumentación y Control con interfaz web interactiva.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-1.31+-red.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

---

## 🚀 Inicio Rápido

### 1. Instalar Dependencias

```bash
pip install streamlit pandas openpyxl plotly xlrd
```

### 2. Ejecutar la Aplicación

```bash
streamlit run app_dashboard_ic.py
```

### 3. Cargar tus Datos

Tienes 3 opciones:

- **🌐 URL desde la nube** (Google Drive, OneDrive, Dropbox)
- **📁 Archivo local** (sube tu archivo Excel)
- **🧪 Generar ejemplo** (archivo de prueba automático)

¡Listo! La aplicación se abrirá en tu navegador en `http://localhost:8501`

---

## 📋 Características Principales

### ✅ Métricas en Tiempo Real
- Visualización de progreso por cada actividad
- Porcentajes de completitud con indicadores visuales
- Contadores automáticos de completados vs pendientes

### 🔍 Filtros Dinámicos Múltiples
- Área, Sistema General, Sistema BMS/SMC/DCS
- Tipo de Instrumento, Prioridad, Hito
- Selección múltiple simultánea
- Aplicación instantánea sin recargar

### ⚠️ Análisis de Pendientes
- Identificación rápida de equipos pendientes por actividad
- Tablas detalladas con información completa
- Exportación individual a CSV por actividad

### 📈 Visualizaciones Interactivas
- Gráficos de barras apiladas (completados vs pendientes)
- Gráficos de porcentaje de completitud con colores
- Gráficos circulares de distribución
- Zoom, pan y descarga de imágenes

### 🔍 Análisis Multidimensional
- **Por Área**: Distribución geográfica de equipos
- **Por Sistema**: Comparación entre sistemas
- **Por Tipo**: Clasificación de instrumentos
- **Por Prioridad**: Equipos críticos

### 📊 Tabla Interactiva Completa
- Vista configurable (simplificada o completa)
- Búsqueda por TAG en tiempo real
- Ordenamiento por columnas
- Exportación en CSV y Excel

### 🚀 Rendimiento Optimizado
- Cache inteligente de 5 minutos
- Soporte para más de 10,000 equipos
- Actualización manual con un clic
- Carga rápida y responsive

---

## 📁 Estructura del Archivo Excel

### Nombre de la Hoja
El archivo Excel debe contener una hoja llamada: **"Equipos I&C"**

### Columnas Requeridas

#### Identificación del Equipo:
- `TAG` - Identificador único del equipo
- `TIPO INSTRUMENTOS` - Tipo de instrumento o equipo
- `DESCRIPTION` - Descripción del equipo
- `AREA` - Área del proyecto
- `SISTEMA GENERAL` - Sistema principal al que pertenece
- `SISTEMA BMS/SMC/DCS` - Sistema de control (DCS, PLC, etc.)
- `SISTEMA` - Subsistema específico
- `SIGNAL ASSOCIATION` - Asociación de señal
- `SIGNAL` - Tipo de señal (4-20mA, Digital, etc.)
- `I/O` - Tipo de entrada/salida (AI, AO, DI, DO)

#### Planificación:
- `Hito` - Hito del proyecto
- `Prioridad` - Prioridad del equipo (Alta, Media, Baja)
- `Pre Emsanblado` - Estado de preensamblado

#### Actividades de Seguimiento:
- `A Instalar` - Listo para instalar
- `Instalación` - Estado de instalación
- `Canalización/Bandeja` - Canalización y bandejas
- `Cableado` - Estado del cableado
- `Conexión Equipo` - Conexión del equipo
- `Conexión DCS` - Conexión al DCS/PLC
- `Marquillado Equipo` - Marquillado del equipo
- `Marquillado Cable` - Marquillado de cables
- `Suiministro de Aire` - Suministro de aire (si aplica)
- `Pre-Comisionamiento` - Precomisionamiento

### Valores para Actividades

**Para indicar "Completado":**
- `OK`
- `SI`
- `Completado`
- `X`
- `1` (número)
- `True`

**Para indicar "Pendiente":**
- Dejar la celda vacía (recomendado)
- `Pendiente`
- `NO`
- `0` (número)
- `False`

---

## 🌐 Configuración de URL desde la Nube

### Google Drive

1. Sube tu archivo Excel a Google Drive
2. Haz clic derecho → **"Obtener enlace"**
3. Configura: **"Cualquiera con el enlace puede ver"**
4. Copia el **FILE_ID** de la URL:
   ```
   https://drive.google.com/file/d/FILE_ID_AQUI/view?usp=sharing
   ```
5. Convierte la URL a formato de descarga:
   ```
   https://drive.google.com/uc?export=download&id=FILE_ID_AQUI
   ```
6. Usa esta URL en la aplicación

### OneDrive

1. Sube tu archivo a OneDrive
2. Haz clic derecho → **"Compartir"**
3. Configura: **"Cualquiera con el vínculo puede ver"**
4. Copia el enlace compartido
5. En la URL, reemplaza:
   - `redir` → `download`
   - `view.aspx` → `download.aspx`
6. Usa la URL modificada en la aplicación

### Dropbox

1. Sube tu archivo a Dropbox
2. Obtén el enlace compartido
3. Cambia el parámetro final de `dl=0` a `dl=1`
4. Usa esta URL en la aplicación

---

## 🎯 Guía de Uso

### Cargar Datos desde la Nube

1. En el panel lateral, ve a la pestaña **"🌐 URL"**
2. Pega la URL de tu archivo Excel
3. Marca la casilla **"Usar URL"**
4. Haz clic en **"🔄 Cargar/Actualizar Datos"**

### Cargar Archivo Local

1. En el panel lateral, ve a la pestaña **"📁 Archivo"**
2. Haz clic en **"Browse files"**
3. Selecciona tu archivo Excel
4. El archivo se carga automáticamente

### Generar Archivo de Ejemplo

1. En el panel lateral, ve a la pestaña **"🧪 Ejemplo"**
2. Haz clic en **"📊 Generar Archivo de Ejemplo"**
3. Se creará un archivo `Equipos_IC_Ejemplo.xlsx` en tu directorio
4. Cárgalo desde la pestaña **"📁 Archivo"** para probarlo

### Aplicar Filtros

1. En el panel lateral, sección **"🔍 Filtros"**
2. Desmarca **"Todas"** o **"Todos"** en el filtro que desees
3. Selecciona los valores específicos
4. Los gráficos y tablas se actualizan automáticamente
5. Puedes combinar múltiples filtros simultáneamente

### Ver Equipos Pendientes

1. Baja a la sección **"⚠️ Equipos Pendientes por Actividad"**
2. Selecciona la actividad en el menú desplegable
3. Verás la lista de equipos pendientes
4. Haz clic en **"📥 Descargar Pendientes"** para exportar a CSV

### Buscar un TAG Específico

1. Ve a la sección **"📋 Tabla Completa de Equipos"**
2. En el campo **"🔍 Buscar por TAG"**, escribe el TAG
3. La búsqueda es en tiempo real (no necesitas presionar Enter)
4. Funciona con búsqueda parcial (ej: "FT" encuentra todos los FT-xxx)

### Exportar Datos

1. Aplica los filtros que necesites
2. Ve a **"📋 Tabla Completa de Equipos"**
3. Elige el formato:
   - **CSV**: Más ligero, compatible con Excel
   - **Excel**: Mantiene formato, ideal para reportes
4. El archivo se descarga con timestamp automático

### Análisis Multidimensional

1. Ve a la sección **"🔍 Análisis Multidimensional"**
2. Haz clic en las pestañas:
   - **📍 Por Área**: Distribución por áreas
   - **⚙️ Por Sistema**: Análisis por sistema
   - **🔧 Por Tipo**: Clasificación por tipo
   - **🎯 Por Prioridad**: Vista de prioridades
3. Cada pestaña muestra gráficos y tablas específicas

---

## 💡 Tips y Trucos

### 🎯 Filtrado Estratégico

**Ejemplo 1: Equipos críticos pendientes**
- Filtro "Prioridad": Selecciona solo "Alta"
- Ve a "Equipos Pendientes" y selecciona la actividad
- Exporta la lista para asignar recursos

**Ejemplo 2: Avance por área**
- Filtro "Área": Selecciona un área específica
- Observa las métricas de avance
- Compara con otras áreas

**Ejemplo 3: Pendientes de DCS**
- Filtro "Sistema BMS/SMC/DCS": Selecciona "DCS"
- Ve a "Equipos Pendientes"
- Selecciona "Conexión DCS"

### 🔍 Búsqueda Eficiente

- **Buscar por prefijo**: Escribe "FT" para todos los transmisores de flujo
- **Buscar por número**: Escribe "001" para equipos que terminen en 001
- **Búsqueda parcial**: "PT-2" encuentra PT-200, PT-201, PT-250, etc.

### 📊 Análisis Rápido

1. **Vista ejecutiva**: Sin filtros, mira las métricas generales
2. **Comparación de áreas**: Usa la pestaña "Por Área"
3. **Cuellos de botella**: Identifica actividades con bajo %
4. **Priorización**: Filtra por "Prioridad Alta" y actividad pendiente

### 🚀 Rendimiento

- El sistema cachea datos por 5 minutos automáticamente
- Si actualizas el Excel en la nube, haz clic en "🔄 Cargar/Actualizar Datos"
- Para análisis de más de 10,000 equipos, usa filtros para segmentar

### 📥 Exportación Inteligente

- Antes de exportar, aplica todos los filtros necesarios
- El nombre del archivo incluye fecha y hora automáticamente
- CSV es más rápido, Excel mantiene mejor el formato
- Exporta por actividad para listas de trabajo específicas

---

## ⚙️ Comandos Útiles

### Instalación Completa

```bash
# Crear entorno virtual (recomendado)
python -m venv venv

# Activar entorno (Windows)
venv\Scripts\activate

# Activar entorno (Linux/Mac)
source venv/bin/activate

# Instalar dependencias
pip install streamlit pandas openpyxl plotly xlrd
```

### Ejecución

```bash
# Ejecución estándar
streamlit run app_dashboard_ic.py

# Especificar puerto
streamlit run app_dashboard_ic.py --server.port 8502

# Modo desarrollo (auto-reload)
streamlit run app_dashboard_ic.py --server.runOnSave true

# Ver en red local
streamlit run app_dashboard_ic.py --server.address 0.0.0.0
```

### Mantenimiento

```bash
# Limpiar caché
streamlit cache clear

# Ver configuración
streamlit config show

# Ver versión
streamlit --version
```

---

## 🐛 Solución de Problemas

### ❌ Error: "No se pudo cargar el archivo"

**Posibles causas:**
- URL incorrecta o no es enlace directo
- Archivo sin permisos públicos
- Hoja no se llama "Equipos I&C"

**Soluciones:**
1. Verifica que la URL sea de descarga directa
2. Asegura permisos públicos de lectura
3. Confirma el nombre exacto de la hoja
4. Prueba cargar el archivo localmente primero

### ❌ Las métricas muestran 0%

**Posibles causas:**
- Columnas de actividades sin datos
- Formato de datos incorrecto

**Soluciones:**
1. Verifica que uses valores como: OK, SI, Completado, X, 1
2. Asegura que los nombres de columnas coincidan exactamente
3. No uses espacios extra en los valores

### ❌ Los filtros no responden

**Soluciones:**
1. Recarga la página (Ctrl + R o F5)
2. Haz clic en "🔄 Cargar/Actualizar Datos"
3. Verifica que haya datos después de aplicar los filtros

### ❌ Gráficos no se muestran

**Soluciones:**
1. Verifica conexión a internet
2. Prueba con otro navegador (Chrome recomendado)
3. Asegura que Plotly esté instalado: `pip install plotly`

### ❌ Streamlit no se reconoce como comando

**Soluciones:**
```bash
# Opción 1: Ejecutar como módulo
python -m streamlit run app_dashboard_ic.py

# Opción 2: Reinstalar Streamlit
pip uninstall streamlit
pip install streamlit

# Opción 3: Verificar PATH
pip show streamlit
```

### ❌ Error de memoria con archivos grandes

**Soluciones:**
1. Usa filtros para reducir datos visualizados
2. Aumenta memoria disponible para Python
3. Cierra otras aplicaciones
4. Considera dividir el archivo en hojas por área/sistema

---

## 🔒 Seguridad y Privacidad

- ✅ **Solo lectura**: La aplicación solo lee el archivo, no modifica datos
- ✅ **Sin credenciales**: No se almacenan contraseñas ni tokens
- ✅ **Cache local**: Datos temporales solo en tu computadora
- ✅ **Sin envío externo**: Ningún dato se envía a servidores externos
- ✅ **Código abierto**: Todo el código es visible y auditable

---

## 📱 Compatibilidad

### Navegadores Soportados
- ✅ Chrome / Chromium (recomendado)
- ✅ Firefox
- ✅ Safari
- ✅ Edge
- ✅ Opera

### Sistemas Operativos
- ✅ Windows 10/11
- ✅ macOS 10.14+
- ✅ Linux (Ubuntu, Debian, Fedora, etc.)

### Versiones de Python
- ✅ Python 3.8
- ✅ Python 3.9 (recomendado)
- ✅ Python 3.10
- ✅ Python 3.11
- ✅ Python 3.12

---

## 📦 Dependencias

| Paquete | Versión Mínima | Uso |
|---------|----------------|-----|
| streamlit | 1.31.0 | Framework web interactivo |
| pandas | 2.0.0 | Procesamiento de datos |
| openpyxl | 3.1.0 | Lectura de archivos Excel |
| plotly | 5.18.0 | Visualizaciones interactivas |
| xlrd | 2.0.1 | Soporte para archivos .xls antiguos |

### Instalación de Dependencias

```bash
# Instalación individual
pip install streamlit
pip install pandas
pip install openpyxl
pip install plotly
pip install xlrd

# O todas a la vez
pip install streamlit pandas openpyxl plotly xlrd
```

---

## 🚀 Despliegue en la Nube

### Streamlit Cloud (Gratis)

1. Sube `app_dashboard_ic.py` a GitHub
2. Crea un archivo `requirements.txt`:
   ```
   streamlit==1.31.0
   pandas==2.2.0
   openpyxl==3.1.2
   plotly==5.18.0
   xlrd==2.0.1
   ```
3. Ve a [share.streamlit.io](https://share.streamlit.io)
4. Conecta tu repositorio
5. Selecciona el archivo `app_dashboard_ic.py`
6. ¡Despliega!

### Heroku

```bash
# Crear Procfile
echo "web: streamlit run app_dashboard_ic.py --server.port $PORT" > Procfile

# Crear setup.sh
echo "mkdir -p ~/.streamlit/" > setup.sh
echo "echo '[server]\nport = $PORT\nenableCORS = false\n' > ~/.streamlit/config.toml" >> setup.sh

# Desplegar
heroku create tu-app-dashboard
git push heroku main
```

---

## 📊 Capacidad del Sistema

- **Equipos soportados**: Más de 10,000
- **Columnas**: Ilimitadas
- **Filtros simultáneos**: Hasta 6
- **Usuarios concurrentes**: Depende del hosting
- **Tiempo de carga**: < 2 segundos (con cache)
- **Tamaño de archivo**: Hasta 50 MB recomendado

---

## 🎓 Casos de Uso Reales

### Para Project Managers
- Dashboard ejecutivo con KPIs principales
- Identificación de retrasos por actividad
- Reportes automáticos para stakeholders
- Seguimiento de hitos del proyecto

### Para Ingenieros de Campo
- Listas de equipos pendientes por actividad
- Filtrado por área de trabajo específica
- Búsqueda rápida de TAGs
- Exportación de listas de trabajo diarias

### Para Coordinadores
- Distribución de carga de trabajo por sistema
- Análisis de recursos por área
- Identificación de cuellos de botella
- Planificación de actividades

### Para Supervisores
- Seguimiento de prioridades
- Control de avance por hito
- Verificación de cumplimiento
- Auditoría de actividades

---

## 📈 Próximas Mejoras (Roadmap)

- [ ] Gráfico de Gantt temporal
- [ ] Exportación a PDF con gráficos
- [ ] Dashboard de tendencias históricas
- [ ] Modo oscuro/claro
- [ ] Alertas automáticas por correo
- [ ] Integración con MS Project
- [ ] API REST para integraciones
- [ ] App móvil complementaria
- [ ] Sistema de comentarios por equipo
- [ ] Control de versiones de archivos

---

## 🤝 Contribuciones

Este proyecto está abierto a mejoras y sugerencias. Si deseas contribuir:

1. Identifica una mejora o corrección
2. Realiza los cambios en el código
3. Documenta los cambios en este README
4. Comparte tu versión mejorada

---

## 📄 Licencia

Este proyecto se distribuye bajo la Licencia MIT. Eres libre de usar, modificar y distribuir este software con atribución apropiada.

---

## 📞 Soporte

Para problemas, dudas o sugerencias:

1. Revisa la sección "Solución de Problemas" arriba
2. Verifica que tu archivo Excel cumpla con la estructura requerida
3. Asegura que todas las dependencias estén instaladas correctamente
4. Prueba con el archivo de ejemplo generado automáticamente

---

## ✨ Créditos

**Dashboard Equipos I&C v1.0.0**

Desarrollado para proyectos de Instrumentación y Control Industrial

**Tecnologías utilizadas:**
- [Streamlit](https://streamlit.io) - Framework web
- [Pandas](https://pandas.pydata.org) - Análisis de datos
- [Plotly](https://plotly.com) - Visualizaciones interactivas
- [OpenPyXL](https://openpyxl.readthedocs.io) - Manejo de Excel

---

## 🎯 Resumen Rápido

```bash
# 1. Instalar
pip install streamlit pandas openpyxl plotly xlrd

# 2. Ejecutar
streamlit run app_dashboard_ic.py

# 3. Cargar datos
# → URL desde la nube, archivo local, o generar ejemplo

# 4. Analizar
# → Métricas, gráficos, filtros, y exportaciones

# ¡Listo! 🎉
```

---

**¡Gracias por usar Dashboard Equipos I&C!** 🏭📊✨
