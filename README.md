# 🧑‍💻 Autor
MIGUEL ANGEL MADARIAGA CANTOR
PRACTICANTE UNIVERSITARIO FESC
CLINICA SAN JOSE DE CUCUTA S.A
Desarrollado como solución de automatización para procesos de liquidación de honorarios médicos.


# 🏥 HONORARIOS FACTURADOS

Sistema de automatización para el cálculo y validación de honorarios médicos a partir de archivos de facturación en Excel, organizado por especialidades.

---

## 🚀 Descripción

Este proyecto permite procesar archivos institucionales y generar automáticamente:

- Hoja **PRE**
- Hoja **PEDIR FAC**
- Hoja **AMARILLO**
- Hoja **ANULADOS**
- Hojas adicionales según especialidad (ej: DUPLICADOS en urología)

Incluye lógica avanzada para:

- Validación de datos
- Aplicación de porcentajes por especialidad
- Identificación de anulados
- Detección de duplicados
- Cálculo automático de honorarios
- Generación de cuadros contables finales

---

## 🧠 Especialidades soportadas

- Ortopedia  
- Urología  
- Cardiología  
- Gastroenterología  

---

## ⚙️ Funcionalidades principales

✔ Procesamiento automático de Excel  
✔ Aplicación de reglas por especialidad  
✔ Manejo de porcentajes dinámicos  
✔ Persistencia de configuración (AppData)  
✔ Interfaz gráfica amigable (Tkinter)  
✔ Generación de reportes listos para contabilidad  
✔ Integración con manual de usuario  

---

## 📊 Lógica del sistema

### 🔹 Cálculo de honorarios
- Basado en:
  - COOSALUD
  - HOSVIREPORT
- Aplicación de porcentajes:
  - 100%
  - 75%
  - 70%
  - 60%
  - 50%

### 🔹 Cuadro resumen (PRE)

- **TOTAL BAJAR HOSV** → suma SALDO  
- **TOTAL COOSALUD** → suma VLR A AUTORIZAR  
- **DIF** → diferencia entre ambos  

---

## 🖥️ Interfaz gráfica

El sistema cuenta con una interfaz que permite:

- Seleccionar archivo de entrada
- Definir archivo de salida
- Elegir especialidad
- Ajustar porcentajes
- Validar archivo
- Visualizar bitácora
- Abrir manual de usuario
- Consultar resumen del proceso

---

## 💾 Configuración persistente

Los porcentajes se guardan automáticamente en: settings.json

Esto permite que:

✔ Los cambios se mantengan al cerrar la aplicación  
✔ Cada usuario tenga su propia configuración  

---

## 📦 Generación del ejecutable

El sistema puede compilarse a `.exe` usando PyInstaller:

```bash
pyinstaller --noconfirm --clean --onedir --windowed \
--name HONORARIOS_FACTURADOS \
--icon=honorarios_app/resources/icono_facturados.ico \
--add-data "honorarios_app/resources/manual_usuario_facturados.pdf;resources" \
main.py

ESTRUCTURA DEL PROYECTO

HONORARIOS/
│
├── honorarios_app/
│   ├── core/              # Lógica principal
│   ├── specialties/       # Reglas por especialidad
│   ├── services/          # Factory de especialidades
│   ├── config/            # Configuración (settings)
│   ├── ui/                # Interfaz gráfica
│   └── resources/         # Manual e íconos
│
├── build/
├── dist/
├── main.py
├── requirements.txt
└── README.md


