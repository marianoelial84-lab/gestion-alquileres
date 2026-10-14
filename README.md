# 🏨 Sistema de Gestión Automatizado para Apart Chivilcoy

> **Trabajo Final Integrador** — Diplomatura en IA Aplicada a Entornos Digitales de Gestión (FCE-UBA)  
> **Autora:** María Noelia López

---

## 📌 ¿Qué hace este proyecto y para qué sirve?

**Apart Chivilcoy** es un emprendimiento de alquileres temporarios compuesto por **5 departamentos** (Depto 2 al Depto 6)

Este proyecto resuelve la fragmentación y desorganización operativa (reserva por múltiples canales como Booking, Airbnb y venta directa) mediante un **sistema integral de gestión administrativa, financiera y de disponibilidad**
Centraliza toda la operación en el ecosistema de Google Workspace, automatizando tareas repetitivas y ofreciendo visualización de métricas en tiempo real

### 🎯 Principales Beneficios:
* **Centralización de datos:** Google Sheets actúa como la "Única Fuente de Verdad"
* **Ahorro de tiempo:** Ingesta automática de correos con reservas y facturas de servicios
* **Control total y ágil:** Visualización de disponibilidad y métricas financieras desde la PC o el celular
* **Cero costos de infraestructura:** Opera 100% serverless en la nube gratuita de Google

---

## 🚀 Principales Funcionalidades

### 1. 📬 Ingesta Automática de Reservas y Servicios (Gmail + Apps Script)
* **Reservas:** Detecta correos de confirmación de Booking y Airbnb, extrae datos clave (huésped, fechas, depto, canal) y los registra evitando duplicados mediante IDs únicos y etiquetas
* **Servicios Públicos:** Analiza correos de facturas (Luz/EDEN, Gas/Camuzzi, Agua/ABSA, Internet/Personal), registrando montos y fechas de vencimiento automáticamente

### 2. 🗓️ Sincronización con Google Calendar
* Agenda automáticamente los Check-ins y Check-outs de reservas con recordatorios previos
* Agenda las fechas de vencimiento de los servicios con notificaciones 1 día antes a las 10:00 AM

### 3. 📄 Emisión Automática de Contratos (Google Docs & Drive)
* Genera contratos de alquiler temporario personalizados en PDF a partir de una plantilla, reemplazando datos del huésped y fechas de la reserva.
* Guarda el contrato en Google Drive y lo envía por correo electrónico al cliente.

### 4. 💸 Control de Comisiones
* Extrae comisiones cobradas por canales externos (ej. Airbnb) y las registra como gastos

### 5. 📊 Dashboard Web Interactivo (HTML + Tailwind CSS + Chart.js)
Accesible desde navegadores móviles o escritorio, incluye:
* **Métricas Clave (KPIs):** Ingresos Totales, Gastos Totales, Ganancia Neta, Noches Ocupadas, Tasa de Ocupación %, ADR (Tarifa Promedio Diaria) y RevPAR
* **Filtros Avanzados:** Por Departamento, Año y Selección Múltiple de Meses
* **Gráficos Dinámicos:** 
  * Evolución mensual (Ingresos vs. Gastos)
  * Ventas apiladas por canal por mes
  * Comparativo de Ingresos vs. Gastos por Departamento
  * Gastos desglosados por categoría
* **Módulo de Carga Manual / Escaneo de Comprobantes:** Carga rápida de facturas mediante OCR e IA (Gemini).

---

## 🛠️ Arquitectura Técnica y Tecnologías Utilizadas

* **Base de Datos Central:** Google Sheets
* **Middleware & Automatización:** Google Apps Script (JavaScript)
* **Inteligencia Artificial:** Google Gemini (procesamiento de comprobantes / OCR / IA Gemini)
* **Frontend Web App:** HTML5, Tailwind CSS, Chart.js, FontAwesome.
* **Integraciones Nativas:** Gmail, Google Calendar, Google Docs, Google Drive

---

## 📂 Estructura del Repositorio

```text
├── CodigoAirBnB.gs         # Lógica de extracción de correo para Airbnb
├── CodigoAirBooking.gs         # Lógica de extracción de correo para Booking
├── CodigoGastos.gs           # Lógica para procesar servicios (EDEN, Camuzzi, ABSA, Personal)
├── CodigoCalendar.gs         # Sincronización de eventos y recordatorios
├── CodigoContratos.gs        # Generación y envío automatizado de contratos PDF
├── Index.html                # Interface Web App (Dashboard interactivo y Formulario de Carga)
└── README.md                 # Documentación del proyecto
