# 🤖 Bot de Finanzas Personales con Telegram y Gemini AI

Un bot inteligente que te ayuda a registrar y organizar tus transacciones financieras a través de Telegram, con integración de Gemini AI para entender lenguaje natural.

## ✨ Características

- 📱 **Integración con Telegram**: Registra transacciones directamente desde tu teléfono
- 🤖 **Lenguaje Natural con Gemini AI**: Entiende frases como "gasté 3500 en el supermercado"
- 📧 **Lectura automática de emails**: Procesa automáticamente transacciones de bancos y plataformas de pago
- 💾 **Almacenamiento en Excel**: Todas tus transacciones organizadas en un archivo `.xlsx`
- 📊 **Resumen financiero**: Visualiza tus ingresos, egresos y balance mensual
- 📋 **Historial de transacciones**: Consulta tus últimas 10 transacciones
- 🏷️ **Clasificación automática**: Categoriza tus gastos automáticamente

## 📋 Requisitos Previos

1. **Cuenta en Telegram**: Descargar la app en tu teléfono
2. **Bot en Telegram**: Crear un bot con [@BotFather](https://t.me/botfather)
3. **Cuenta de Gmail**: Con contraseña de aplicación habilitada
4. **API Key de Google Gemini**: Obtener en [Google AI Studio](https://makersuite.google.com/app/apikeys)
5. **Python 3.9+**: Instalado en tu máquina

## 🚀 Instalación y Configuración

### 1. Clonar o descargar el proyecto

```bash
cd /home/kashu/Documentos/BotFinanzas
```

### 2. Crear entorno virtual e instalar dependencias

```bash
python3 -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configurar variables de entorno

Crear un archivo `.env` en la carpeta del proyecto o configurar las variables:

```bash
export TELEGRAM_TOKEN="tu_token_aqui"
export TELEGRAM_CHAT_ID="tu_chat_id"
export GMAIL_USER="tu_email@gmail.com"
export GMAIL_PASS="tu_contraseña_aplicacion"
export GEMINI_API_KEY="tu_api_key_de_gemini"
```

**En Windows (PowerShell):**
```powershell
$env:TELEGRAM_TOKEN = "tu_token_aqui"
$env:GMAIL_USER = "tu_email@gmail.com"
# ... etc
```

### 4. Obtener tus credenciales

#### 📱 Token de Telegram
1. Escribe a [@BotFather](https://t.me/botfather)
2. Comando: `/newbot`
3. Sigue las instrucciones y obtendrás un token como: `123456789:ABCdefGHIjklMNOpqrSTUvwxyz...`

#### 📧 Contraseña de Aplicación de Gmail
1. Ve a [Google Account Security](https://myaccount.google.com/security)
2. Activa "Verificación en dos pasos" (si no lo está)
3. Genera una "Contraseña de aplicación" para "Correo" y "Windows/Mac/Linux"
4. Copia la contraseña (sin espacios): `taaj nicd ecst jpbh` → `taajnicdecstjpbh`

#### 🤖 API Key de Gemini
1. Ve a [Google AI Studio](https://makersuite.google.com/app/apikeys)
2. Haz click en "Create API Key"
3. Elige "Create new secret key in new project"
4. Copia tu API key

### 5. Tu Chat ID de Telegram
1. Escribe a [@userinfobot](https://t.me/userinfobot)
2. Recibirás tu ID: `8465318300` (por ejemplo)

### 6. Preparar archivo Excel

El archivo `registro_finanzas.xlsx` debe tener una hoja llamada "Transacciones" con estas columnas:

| Fecha | Tipo | Categoría | Monto | Descripción | Origen | ID |
|-------|------|-----------|-------|-------------|--------|-----|

Si no existe, el bot lo creará automáticamente.

### 7. Ejecutar el bot

```bash
source venv/bin/activate
python bot_finanzas.py
```

Deberías ver:
```
🤖 Bot iniciado correctamente y listo para operar...
🔗 Conectado a Telegram: ✅
🤖 Gemini API: ✅
📧 Email: ✅
```

## 💬 Cómo Usar el Bot

### Comandos disponibles

| Comando | Descripción |
|---------|-------------|
| `/start` | Ver ayuda y opciones |
| `/ayuda` | Mostrar ayuda nuevamente |
| `/revisar_emails` | Fuerza la búsqueda de correos sin leer |
| `/ultimas` | Muestra las últimas 10 transacciones |
| `/resumen` | Resumen financiero del mes actual |

### Ejemplos de Mensajes Naturales

El bot entiende estos tipos de mensajes gracias a Gemini AI:

```
Gasté 3500 pesos en el supermercado
Compré en Coto $2800
Me llegaron 50000 de freelance
+15000 por vender mis libros
Retiro 5 mil en efectivo
Pague 1200 de luz
Transferencia de 10000 a Juan
Ingreso por trabajo remoto $25000
```

### Formato Tradicional (sin Gemini)

Si no tienes Gemini configurado, también funciona con:

```
3500 retiro de dinero
+50000 cobro de proyecto
2800 supermercado
```

## 📁 Estructura del Proyecto

```
BotFinanzas/
├── bot_finanzas.py          # Archivo principal del bot
├── registro_finanzas.xlsx   # Base de datos de transacciones
├── requirements.txt         # Dependencias Python
├── .env                     # Variables de entorno (crear)
└── README.md               # Este archivo
```

## 🔧 Solución de Problemas

### Error: "Token inválido"
- Verifica que el TELEGRAM_TOKEN sea correcto
- No incluyas espacios extras

### Error: "Autenticación fallida en Gmail"
- Asegúrate de usar la "Contraseña de aplicación", no tu contraseña regular
- Quita guiones de la contraseña: `taaj-nicd-ecst-jpbh` → `taajnicdecstjpbh`
- Verifica que la verificación en dos pasos esté activa

### Error: "Gemini API Key inválida"
- Verifica la clave en https://makersuite.google.com/app/apikeys
- Asegúrate de que no haya espacios extras
- Revisa que sea una clave de secret, no de public

### El bot no recibe mensajes
- Asegúrate de escribirle al bot en Telegram
- Comprueba que el bot está corriendo
- Usa `/start` para iniciar la conversación

## 🛡️ Seguridad

⚠️ **IMPORTANTE**: 
- **NUNCA** compartas tu TELEGRAM_TOKEN o GEMINI_API_KEY
- **NUNCA** commiteés credenciales en un repositorio git
- Usa variables de entorno o archivos `.env` (que no estén en git)

## 📊 Categorías Disponibles

El bot clasifica automáticamente en:

- Supermercado / Comida
- Salidas / Entretenimiento
- Servicios / Impuestos
- Transferencia Enviada / Recibida
- Sueldo / Ingreso
- Retiro de Efectivo
- Transporte / Combustible
- Salud
- Entretenimiento
- Comidas Afuera
- Compras Varias
- Otros Ingresos

## 🚀 Próximas Mejoras

- [ ] Exportar reportes en PDF
- [ ] Gráficos de gastos mensuales
- [ ] Alertas automáticas de gastos altos
- [ ] Base de datos en lugar de Excel
- [ ] Múltiples usuarios
- [ ] Sincronización en la nube

## 📝 Licencia

Este proyecto es de uso personal. Siéntete libre de modificarlo.

## 👨‍💻 Autor

Creado con ❤️ para facilitar el control financiero personal.

---

**¿Preguntas?** Revisa el código o contacta al desarrollador.
