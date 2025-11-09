# Robot de Firma de Justificaciones

Sistema de automatización RPA (Robotic Process Automation) para la firma digital de justificaciones en el portal Red.es, utilizando certificados digitales personales.

## 📋 Descripción

Este robot automatiza el proceso de firma de justificaciones pendientes en el portal de gestión de Red.es. Utiliza Playwright para la automatización web y maneja automáticamente la selección de certificados digitales, eliminando la necesidad de intervención manual en los diálogos del sistema operativo.

### Características principales

- ✅ **Selección automática de certificado**: Configura dinámicamente el navegador para auto-seleccionar el certificado elegido
- ✅ **Múltiples estrategias de firma**: Prioriza firma con Cl@ve, fallback a AutoFirma
- ✅ **Interfaz web intuitiva**: Control completo desde el navegador
- ✅ **Procesamiento por lotes**: Firma automáticamente todos los expedientes pendientes
- ✅ **Multi-certificado**: Soporta diferentes certificados para diferentes usuarios
- ✅ **Logs en tiempo real**: Seguimiento detallado del proceso

## 🛠️ Requisitos

- **Python 3.8+**
- **Windows** (requerido para pywinauto y manejo de certificados)
- **Google Chrome** instalado
- **Certificado digital** instalado en Windows (Almacén personal del usuario)
- **Acceso** al portal Red.es

## 📦 Instalación

### 1. Clonar el repositorio

```bash
cd "c:\Users\TuUsuario\Documents\04. ROBOTS\04. ROBOT FIRMA JUSTIFICACIONES"
```

### 2. Instalar dependencias

```bash
cd just-signer
pip install -r requirements.txt
playwright install chromium
```

### 3. Verificar instalación de certificados

Asegúrate de que tu certificado digital está instalado:
1. Presiona `Win + R` y escribe `certmgr.msc`
2. Ve a **Personal** → **Certificados**
3. Verifica que tu certificado aparece en la lista

## 🚀 Uso

### Iniciar el robot

```bash
cd just-signer
python app.py
```

Verás un mensaje como:
```
UI disponible en http://localhost:8771 | ENGINE=async-only-1
```

### Usar la interfaz web

1. **Abrir el navegador** y ve a: http://localhost:8771

2. **Seleccionar certificado**: El robot detectará automáticamente todos los certificados válidos instalados en tu sistema y los mostrará en un selector desplegable.

3. **Configurar opciones**:
   - **Categoría**: Kit Digital (KD) o Kit Consulting (KC)
   - **Velocidad**: Rápido, Medio o Lento

4. **Iniciar proceso**: Pulsa "Iniciar Proceso"

5. **Monitorear**: El robot abrirá Chrome y comenzará la automatización. Podrás ver los logs en tiempo real en la interfaz web.

### Flujo de trabajo del robot

1. **Apertura del portal** → Navega a Red.es
2. **Autenticación en Cl@ve** → Selecciona automáticamente el certificado configurado
3. **Búsqueda de expedientes** → Localiza todos los expedientes "Pdte. presentar"
4. **Firma de expedientes**:
   - **Prioridad 1**: Firma con Cl@ve (integrada, sin diálogos nativos)
   - **Fallback**: AutoFirma (si Cl@ve no está disponible)
5. **Iteración**: Repite para todos los expedientes pendientes en todas las páginas
6. **Finalización**: Cierra el navegador y muestra el resumen

## 🏗️ Arquitectura

### Componentes principales

```
just-signer/
├── app.py              # Servidor Flask + SocketIO (API REST + WebSockets)
├── robot_async.py      # Motor de automatización Playwright (async)
├── requirements.txt    # Dependencias Python
├── templates/
│   └── index.html      # Interfaz de usuario web
└── tools/
    └── cert_clicker.py # Helper para diálogos nativos de Windows
```

### Tecnologías utilizadas

- **Flask + Flask-SocketIO**: Servidor web y comunicación en tiempo real
- **Playwright (async)**: Automatización del navegador
- **pywinauto**: Interacción con diálogos nativos de Windows (UI Automation)
- **PowerShell**: Extracción de certificados del sistema Windows

## 🔐 Manejo de Certificados Digitales

### Política de auto-selección

El robot implementa la política `AutoSelectCertificateForUrls` de Chromium de forma **dinámica**:

```python
# Configuración dinámica al iniciar Chrome
--auto-select-certificate-for-urls=[
  {
    "pattern": "https://pasarela.clave.gob.es",
    "filter": {
      "SUBJECT": {"CN": "NOMBRE APELLIDO - DNI"},
      "ISSUER": {"CN": "AC FNMT Usuarios"}
    }
  }
]
```

### Ventajas de esta implementación

✅ **Sin permisos de administrador**: No modifica el registro de Windows  
✅ **Multi-usuario**: Cada ejecución usa el certificado seleccionado por el usuario  
✅ **Dinámico**: No requiere configuración previa del sistema  
✅ **Simultáneo**: Múltiples instancias pueden usar certificados diferentes  

### Estrategias de fallback

Si la auto-selección por política falla, el robot implementa tres niveles de fallback:

1. **Helper local** (`cert_clicker.py`): Proceso independiente que detecta y cierra el diálogo nativo
2. **UI Automation directa** (pywinauto): Interacción con el diálogo desde el hilo principal
3. **Selección en DOM** (Playwright): Si el diálogo aparece en la interfaz web de Cl@ve

## 🎯 Configuración Avanzada

### Velocidades de ejecución

```python
DELAY_PRESETS = {
    "rapido": 0.25,   # Para sistemas rápidos y conexiones estables
    "medio": 0.6,     # Recomendado para uso general
    "lento": 1.2,     # Para sistemas lentos o conexiones inestables
}
```

### Puerto del servidor

Para cambiar el puerto (por defecto 8771), edita `app.py`:

```python
PORT = 8771  # Cambia a tu puerto deseado
```

### Modo headless

Para ejecutar el navegador sin interfaz gráfica (testing), edita `app.py`:

```python
HEADLESS = True  # False para ver el navegador
```

## 🐛 Resolución de Problemas

### El robot no encuentra el certificado

1. Verifica que el certificado está en `certmgr.msc` → Personal → Certificados
2. Asegúrate de que el certificado no ha caducado
3. Comprueba que el emisor es "AC FNMT Usuarios" u otro reconocido

### El diálogo de certificado sigue apareciendo

El robot tiene múltiples estrategias de fallback. Si ves el diálogo:
- El robot intentará cerrarlo automáticamente con `cert_clicker.py`
- Si falla, usará UI Automation (pywinauto)
- Revisa los logs para ver qué estrategia se está utilizando

### Error "pywinauto not found"

```bash
pip install pywinauto>=0.6.8
```

### Error de Playwright

```bash
playwright install chromium
```

### El spinner bloquea el botón de firma

El robot ahora espera automáticamente a que desaparezca el spinner antes de hacer clic. Si el problema persiste, aumenta el timeout en `robot_async.py`:

```python
await spinner.wait_for(state="hidden", timeout=10000)  # Aumenta a 15000 o 20000
```

## 📝 Logs

Los logs se muestran en tiempo real en la interfaz web. Formato típico:

```
[23:32:33] [Robot] Velocidad: medio (delay 0.6s)
[23:32:33] [Robot] AutoSelectCertificateForUrls (SUBJECT='NOMBRE APELLIDO - DNI' ISSUER='AC FNMT Usuarios')
[23:32:35] [Robot] Navegando a Justificaciones (Kit Digital)
[23:32:39] [Robot] Autenticación requerida en Cl@ve
[23:32:44] [Robot] Certificado seleccionado por diálogo nativo (UIA)
[23:32:48] [Robot] Autenticado en Cl@ve
[23:32:48] [Robot] Buscando expedientes 'Pdte. presentar' en la página actual...
[23:32:48] [Robot] Abriendo expediente KD/0001234567-001 (Estado: Pdte. presentar)
[23:32:49] [Robot] Intentando firmar expediente (Preferencia: Cl@ve, Fallback: AutoFirma)...
[23:32:50] [Robot] Pulsado 'Firma con Cl@ve y presentar'.
[23:32:51] [Robot] Pasarela Cl@ve detectada, seleccionando certificado...
[23:32:52] [Robot] Certificado seleccionado por helper local y aceptado (rápido).
[23:32:54] [Robot] Expediente firmado. Volviendo al listado...
```

## 🔄 Actualizaciones

Para actualizar las dependencias:

```bash
pip install --upgrade -r requirements.txt
```

## 📄 Licencia

Este proyecto es de uso interno. Todos los derechos reservados.

## 👥 Soporte

Para reportar problemas o sugerencias, contacta con el equipo de desarrollo.

---

**Versión**: 1.0.0  
**Motor**: async-only-1  
**Última actualización**: Enero 2025
