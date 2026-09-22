# FASE 1 - Levantamiento funcional y análisis completo del sistema

## Sistema de Automatización de Honorarios por lo Facturado

**Fecha del análisis:** 19 de junio de 2026  
**Alcance:** código fuente, interfaz Tkinter, configuración, reglas por especialidad, pruebas, empaquetado, recursos PDF/ICO y artefactos de salida.  
**Criterio:** este documento describe únicamente comportamientos comprobables en el proyecto. Cuando una capacidad está incompleta, sin uso o presenta una diferencia entre interfaz, documentación y ejecución, se indica expresamente.

---

# 1. Descripción general del sistema

## 1.1 Propósito

El aplicativo automatiza la preparación y liquidación de honorarios médicos facturados a partir de un libro Excel institucional. Lee datos de facturación, tarifas COOSALUD y, según la especialidad, datos HOSVIREPORT; aplica exclusiones, anulaciones, duplicados y cálculos de honorarios; y genera un nuevo libro Excel organizado para revisión administrativa.

## 1.2 Alcance funcional comprobado

El sistema implementa:

- interfaz gráfica de escritorio para Windows mediante Tkinter;
- ejecución alternativa por línea de comandos;
- cuatro especialidades: Ortopedia, Urología, Cardiología y Gastroenterología;
- porcentajes configurables por especialidad y persistencia en el perfil local del usuario;
- validación preliminar de existencia de hojas requeridas;
- detección dinámica de encabezados y columnas mediante nombres normalizados;
- reglas comunes de exclusión;
- reglas de anulados por especialidad;
- reglas de duplicados para Ortopedia, Urología y Cardiología;
- cálculo con tarifas COOSALUD y HOSVIREPORT cuando corresponde;
- generación de PRE, PEDIR FAC, AMARILLO, ANULADOS y, cuando aplica, DUPLICADOS;
- bitácora visual, progreso, resumen de resultados y apertura del archivo generado;
- distribución como aplicación Windows mediante PyInstaller.

No existen módulos funcionales adicionales distintos de “Facturados”. `module_factory.py`, `config_service.py` y `validators.py` están vacíos.

## 1.3 Usuarios objetivo

Por el propósito declarado en la interfaz, README y PDF incluido, los usuarios objetivo son personal administrativo y contable encargado de revisar facturación y liquidar honorarios médicos. El código no implementa autenticación, perfiles, permisos ni diferenciación de roles.

## 1.4 Flujo general

1. Iniciar la aplicación gráfica o la CLI.
2. Seleccionar especialidad.
3. Revisar o modificar porcentajes.
4. seleccionar el Excel de entrada;
5. aceptar o definir la ruta de salida;
6. opcionalmente validar la presencia de hojas requeridas;
7. procesar el libro;
8. clasificar filas con prioridad: ANULADOS, duplicados previos cuando aplica, exclusiones, cálculo y duplicados posteriores cuando aplica;
9. generar y guardar el Excel final;
10. revisar métricas, bitácora, resumen y archivo de salida;
11. decidir si un porcentaje modificado se conserva como nuevo valor base.

---

# 2. Inventario completo de módulos

| Módulo o recurso | Archivo principal | Función | Dependencias principales | Observaciones |
|---|---|---|---|---|
| Entrada gráfica | `main.py` | Invoca `run_app()` | `ui.main_window` | Punto de entrada del ejecutable. |
| Entrada CLI | `honorarios_app/cli.py` | Recibe rutas, especialidad y nombres de hojas; ejecuta `process_excel()` | `argparse`, `core.processor` | No está enlazada desde la GUI. Los errores no se convierten en mensajes amigables; se propagan a consola. |
| Ventana principal | `honorarios_app/ui/main_window.py` | Construye toda la interfaz, eventos, diálogos, progreso y ejecución | Tkinter, configuración, procesador, SO Windows | Única ventana propia del sistema; los demás “pantallas” son diálogos del sistema o messageboxes. |
| Procesador central | `honorarios_app/core/processor.py` | Orquesta lectura, validaciones, reglas, cálculo, hojas y guardado | OpenPyXL, utilidades Excel, especialidades, escritor de salida | Flujo compartido para las cuatro especialidades. |
| Utilidades comunes | `honorarios_app/core/common.py` | Normalización, números, porcentajes y búsqueda de encabezados/columnas | `re` | Los encabezados se buscan en las primeras 200 filas. |
| Utilidades Excel | `honorarios_app/core/excel_utils.py` | Carga libros, construye tarifario, HOSVI y override de porcentajes | OpenPyXL, `common` | `validate_required_sheets()` existe pero el procesador no lo llama. |
| Escritura y formato | `honorarios_app/core/output_writer.py` | Estilos, autoajuste y cuadros resumen | OpenPyXL | `add_totals_row()` existe pero no se usa. |
| Modelo de resultado | `honorarios_app/core/models.py` | DTO `ProcessResult` | `dataclasses` | Devuelve ruta y conteos de cuatro hojas. |
| Validadores | `honorarios_app/core/validators.py` | Sin implementación | Ninguna | Archivo vacío; la validación real está distribuida en UI, procesador y especialidades. |
| Contrato base | `honorarios_app/specialties/base.py` | Reglas comunes y métodos abstractos | `common` | Centraliza exclusiones comunes. |
| Ortopedia | `honorarios_app/specialties/ortopedia.py` | Anulados, duplicados y cálculo ortopédico | Base, HOSVI, COOSALUD | Duplicados sincronizados entre PRE/PEDIR FAC. |
| Urología | `honorarios_app/specialties/urologia.py` | Anulados, duplicados y cálculo urológico | Base, HOSVI, COOSALUD | Tiene distribución especial para porcentajes 1.75 y 1.05. |
| Cardiología | `honorarios_app/specialties/cardiologia.py` | Anulados, duplicados y cálculo cardiológico | Base, COOSALUD | No consume HOSVIREPORT. Maneja dos porcentajes. |
| Gastroenterología | `honorarios_app/specialties/gastroenterologia.py` | Anulados y cálculo gastroenterológico | Base, HOSVI, COOSALUD | No implementa hoja DUPLICADOS. |
| Fábrica de especialidades | `honorarios_app/services/specialty_factory.py` | Convierte nombre en clase de reglas | Cuatro especialidades | Rechaza nombres no soportados. |
| Fábrica de módulos | `honorarios_app/services/module_factory.py` | Sin implementación | Ninguna | Archivo vacío. |
| Servicio de configuración | `honorarios_app/services/config_service.py` | Sin implementación | Ninguna | Archivo vacío. |
| Configuración activa | `honorarios_app/config/settings.py` | Valores por defecto, lectura y persistencia por usuario | `json`, `os` | Guarda en `%LOCALAPPDATA%/HONORARIOS/settings.json`; si no existe, lo crea. |
| JSON del repositorio | `honorarios_app/config/settings.json` | Contiene valores por defecto | Ninguna | No es leído por `settings.py`; funciona como referencia, no como fuente de ejecución. |
| Pruebas | `tests/test_smoke.py` | Casos de anulados y duplicados para Cardiología/Ortopedia | OpenPyXL, procesador, reglas | Cobertura parcial: no cubre UI, Urología ni Gastroenterología. |
| Manual histórico | `honorarios_app/resources/manual_usuario_facturados.pdf` | Manual PDF accesible desde UI | Visor PDF del sistema | 14 páginas; incluye capturas, pero está desactualizado frente al código actual. |
| Icono | `honorarios_app/resources/icono_facturados.ico` | Icono de ejecutable | PyInstaller | Imagen 68×68 de formulario/gráfica en verde. No se configura explícitamente con `iconbitmap()` en Tkinter. |
| Empaquetado | `HONORARIOS_FACTURADOS.spec` | Construye distribución Windows `onedir`, sin consola | PyInstaller, PDF, ICO | Incluye el PDF como recurso y usa el ICO como icono del EXE. |
| Dependencias | `requirements.txt` | Declara `openpyxl` | pip | Tkinter pertenece a Python; PyInstaller no está declarado. |
| README | `README.md` | Descripción técnica y comando de empaquetado | Ninguna | Algunas afirmaciones son generales y no sustituyen el comportamiento del código. |
| Marcadores de paquete | Archivos `__init__.py` | Declaran paquetes Python | Ninguna | No contienen lógica funcional. |
| Artefactos compilados | `build/` y `dist/` | Resultado de PyInstaller | Runtime Python y librerías | Son artefactos generados, no fuente funcional. |

---

# 3. Inventario de todas las pantallas

## 3.1 Ventana principal “Honorarios Médicos Por Facturado”

**Propósito:** concentrar configuración, ejecución, seguimiento y acceso a soporte.

**Estructura:**

- encabezado con título y subtítulo;
- cuatro tarjetas de métricas: Especialidad, Registros PRE, Registros PEDIR FAC, Estado del módulo;
- tarjeta “Parámetros de procesamiento”;
- tarjeta lateral “Soporte y consulta” y “Herramientas”;
- tarjeta “Bitácora del sistema”;
- barra de progreso y texto de estado;
- canvas desplazable con scrollbar vertical.

**Comportamiento adaptable:** por debajo de 1280 px, formulario, acciones y bitácora se apilan; en anchura mayor, formulario y acciones comparten la primera fila y la bitácora ocupa el ancho inferior. Tamaño inicial 1360×810, mínimo 1000×620 e intento de maximización.

**Eventos:** selección de especialidad, cambio de ruta de salida, botones y rueda del ratón.

## 3.2 Diálogo nativo “Seleccionar archivo Excel de entrada”

Abierto por “Buscar archivo”. Acepta `*.xlsx`, `*.xlsm`, `*.xltx` y `*.xltm`. Al elegir archivo llena la entrada y, si la salida sigue automática, propone el nombre de salida.

## 3.3 Diálogo nativo “Definir archivo de salida”

Abierto por “Guardar como”. Solo ofrece `*.xlsx`; establece salida manual y desactiva la actualización automática del nombre hasta limpiar la ruta.

## 3.4 Visor externo del manual

No es una ventana propia. “Manual del módulo” resuelve el PDF dentro de recursos y lo abre con la aplicación predeterminada de Windows.

## 3.5 Explorador de Windows / aplicación Excel

No son pantallas propias. Se invocan para abrir ubicación, seleccionar el archivo generado o abrir el último archivo.

## 3.6 Ventanas emergentes internas

Son `messagebox` de Tkinter:

- ayuda de uso;
- guía del módulo;
- acerca del sistema;
- resumen del último proceso;
- validación correcta/incompleta/error;
- procesamiento completado/interrumpido;
- confirmación para persistir porcentajes;
- avisos por ruta, archivo, manual o resumen no disponibles;
- confirmaciones de resumen copiado y porcentajes restablecidos.

No existen ventanas secundarias personalizadas, pestañas internas, menús, asistentes por pasos ni pantalla de autenticación.

---

# 4. Inventario completo de botones

| Botón | Acción | Implementación | Evento | Resultado esperado |
|---|---|---|---|---|
| Buscar archivo | Abre selector de entrada | `select_input()` | `command` | Carga ruta y sugiere salida automática. |
| Guardar como | Abre selector de salida | `select_output()` | `command` | Define `.xlsx` manual y desactiva modo automático. |
| Procesar archivo | Ejecuta flujo completo | `process()` | `command` | Genera libro, métricas, resumen, log y posible pregunta de persistencia. |
| Limpiar formulario | Restablece sesión visible | `clear_fields()` | `command` | Vuelve a Ortopedia, limpia rutas, progreso, resultado y métricas; conserva log previo y agrega evento. |
| Abrir ubicación del archivo | Abre carpeta o selecciona salida | `open_output_folder()` | `command` | Usa último resultado o ruta definida. |
| Ayuda de uso | Muestra ocho pasos recomendados | `show_help()` | `command` | Messagebox informativo. |
| Guía del módulo | Describe alcance y especialidades | `show_info()` | `command` | Messagebox informativo. |
| Manual del módulo | Abre PDF incluido | `open_manual_pdf()` | `command` | Visor externo o error. |
| Acerca del sistema | Muestra versión de desarrollo y tecnologías | `show_about()` | `command` | Messagebox informativo. |
| Ver bitácora | Enfoca y desplaza el log al final | `focus_log()` | `command` | Foco en área de texto y nuevo evento de log. |
| Resumen del proceso | Muestra último resumen | `quick_report()` | `command` | Messagebox con ruta, especialidad, conteos y porcentajes. |
| Validar archivo | Comprueba accesibilidad y hojas | `validate_file()` | `command` | Éxito, advertencia de hojas faltantes o error técnico. |
| Abrir último generado | Abre archivo de la sesión | `open_last_generated()` | `command` | Abre Excel/aplicación asociada o informa indisponibilidad. |
| Restablecer porcentajes | Recarga valores persistidos | `reset_percentages()` | `command` | Revierte cambios visibles no guardados. |
| Copiar resumen | Copia resumen al portapapeles | `copy_summary()` | `command` | Portapapeles actualizado o aviso sin proceso. |

Todos estos botones se deshabilitan durante el procesamiento.

---

# 5. Inventario de controles

## 5.1 Labels

- título y subtítulo;
- títulos de tarjetas;
- etiquetas de métricas;
- Especialidad;
- Valor % base / Valor % a actualizar;
- cuatro etiquetas específicas de porcentajes cardiológicos;
- Archivo de entrada / Archivo de salida;
- texto de ayuda sobre nombre automático y porcentajes;
- texto de estado bajo la barra de progreso.

## 5.2 Combobox

`combo_specialty`, solo lectura, con valores exactos:

- `ortopedia` (predeterminado);
- `urologia`;
- `cardiologia`;
- `gastroenterologia`.

Al cambiar, actualiza tarjeta de especialidad, porcentajes visibles, modo de campos y nombre automático de salida.

## 5.3 Textbox / Entry

- porcentaje base simple: solo lectura;
- porcentaje actualizado simple: editable para Ortopedia, Urología y Gastroenterología;
- porcentaje base consultas/cuidados: solo lectura, visible en Cardiología;
- porcentaje actualizado consultas/cuidados: editable;
- porcentaje base procedimientos: solo lectura;
- porcentaje actualizado procedimientos: editable;
- ruta de entrada: editable, además de selector;
- ruta de salida: editable, además de selector;
- bitácora: `tk.Text`, visualmente editable porque no se configura como solo lectura.

## 5.4 CheckBox

No existe ningún CheckBox implementado.

## 5.5 Barras y desplazamiento

- barra horizontal de progreso 0–100;
- scrollbar vertical general de la ventana;
- scrollbar vertical propio de la bitácora;
- soporte de rueda de ratón Windows/Linux.

## 5.6 Tarjetas métricas

- Especialidad;
- Registros PRE;
- Registros PEDIR FAC;
- Estado: LISTO, PROCESANDO, FINALIZADO o ERROR.

## 5.7 Mensajes

Tres categorías: informativos, advertencias y errores; además, confirmación Sí/No para guardar porcentajes base. El detalle completo se presenta en las secciones 8 y 9.

## 5.8 Íconos

- `icono_facturados.ico`: usado por PyInstaller como icono del ejecutable;
- los messageboxes usan iconografía estándar de Tkinter/Windows;
- la ventana no establece icono propio mediante código Tkinter.

---

# 6. Flujo funcional completo

## 6.1 Inicio y carga de recursos

1. `main.py` llama `run_app()`.
2. Tkinter crea `MainWindow`.
3. Se inicializan variables: Ortopedia, porcentajes 70/90/36, estado listo y métricas cero.
4. Se construyen estilos, ventana desplazable, tarjetas, controles y log.
5. Se cargan porcentajes persistidos desde `%LOCALAPPDATA%/HONORARIOS/settings.json`; si no existe, se crea con valores predeterminados.
6. El PDF solo se resuelve cuando el usuario pulsa “Manual del módulo”.
7. El icono se utiliza en el ejecutable compilado, no se carga explícitamente en `MainWindow`.

## 6.2 Selección de especialidad y parámetros

1. La especialidad predeterminada es Ortopedia.
2. Cardiología muestra dos porcentajes; las demás muestran uno.
3. Los valores base son solo lectura; los valores “a actualizar” son editables.
4. Los porcentajes aceptan `%`, coma o punto; deben ser mayores que 0 y menores o iguales a 100.

## 6.3 Selección de archivos

1. El usuario selecciona o escribe la entrada.
2. El sistema propone `PEDIR FAC_<ESPECIALIDAD>_<NOMBRE_ORIGINAL>.xlsx` en la misma carpeta.
3. El usuario puede elegir otra salida; desde ese momento la salida deja de cambiar automáticamente.

## 6.4 Validación previa opcional

La GUI abre el archivo en modo solo lectura y comprueba hojas:

- Cardiología: `COOSALUD`;
- demás: `COOSALUD` y `HOSVIREPORT`.

Esta validación no inspecciona encabezados, columnas, tarifas ni contenido, aunque el mensaje de éxito habla de “estructura mínima compatible”. Las validaciones profundas ocurren únicamente al procesar.

## 6.5 Procesamiento central

1. Valida rutas no vacías y porcentajes.
2. Deshabilita controles.
3. Obtiene la clase de especialidad.
4. Abre el libro con `data_only=True`.
5. Usa la primera hoja como origen, salvo que la CLI indique `--original-sheet`.
6. Exige hojas obligatorias.
7. Lee `PORCENTAJE_PAGO` si existe.
8. Construye tarifario COOSALUD.
9. Para especialidades no cardiológicas, construye índices HOSVIREPORT.
10. Detecta encabezado de la hoja origen buscando, dentro de las primeras 200 filas, textos que contengan `CONTR`, `FAC`, `PROCED` y `COD`.
11. Localiza columnas por encabezados normalizados.
12. Crea libro de salida y hojas.
13. Detecta ANULADOS antes del procesamiento fila a fila.
14. En Urología también calcula duplicados sobre el origen.
15. Recorre filas:
    - ANULADOS tiene prioridad y corta el flujo;
    - duplicado previo de Urología corta el flujo;
    - filas totalmente vacías en factura/contrato/código/procedimiento se omiten;
    - exclusiones comunes pasan a AMARILLO;
    - las demás se calculan y se guardan temporalmente en listas PRE/PEDIR FAC.
16. Ortopedia y Cardiología filtran consultas duplicadas sobre las listas válidas y aplican los mismos índices a PRE y PEDIR FAC.
17. Escribe PRE, DUPLICADOS y PEDIR FAC.
18. Agrega cuadros resumen y fórmulas.
19. Congela panel en A2 y autoajusta columnas (máximo 60 caracteres; inspecciona hasta 2.000 filas).
20. Guarda el `.xlsx`.

## 6.6 Progreso comunicado

- 2%: Abriendo Excel.
- 8%: Leyendo tarifas COOSALUD.
- 14%: Leyendo HOSVIREPORT cuando aplica.
- 18%: Detectando encabezados.
- 22%: Creando archivo de salida.
- 24%: Detectando anulaciones.
- 25%–85%: Procesando filas; actualiza cada 150 filas.
- 88%: Escribiendo hojas y totales.
- 96%: Guardando.
- 100%: Listo.

Los errores del callback de progreso son ignorados para no interrumpir el cálculo.

## 6.7 Finalización

1. Se actualizan métricas y estado.
2. Se muestra resumen con ruta, conteos y porcentajes.
3. Si los porcentajes difieren del valor base, se pregunta si deben persistirse.
4. El usuario puede abrir carpeta, archivo, resumen o copiarlo.
5. Los controles vuelven a habilitarse incluso si ocurre un error.

---

# 7. Especialidades implementadas

## 7.1 Reglas comunes a todas

Contratos de cálculo: `COOSALUD00225` y `COOSALUD00125`. Una fila de otro contrato puede permanecer en PRE/PEDIR FAC, pero su especialidad marca “NO APLICA…” y no actualiza el valor calculado.

Exclusiones a AMARILLO:

1. TIPO FAC igual a `17`;
2. identificación de paciente vacía;
3. nombre de paciente vacío;
4. NO FAC que empieza por `999`;
5. NO FAC igual a `7`;
6. contrato `PARTICULAR525`;
7. código de procedimiento terminado en `-F`.

La razón queda en `RAZON EXCLUSION`.

## 7.2 Ortopedia

**Archivo:** `specialties/ortopedia.py`  
**Entrada:** hoja principal, COOSALUD, HOSVIREPORT; PORCENTAJE_PAGO opcional.  
**Porcentaje base:** 70% clínica.

**Anulados:**

- primera pasada: misma factura + paciente + procedimiento y saldo neto con tolerancia menor a 0,01;
- segunda pasada únicamente sobre filas no anuladas: paciente + procedimiento, ignorando factura, con el mismo neto;
- las filas resultantes van exclusivamente a ANULADOS antes de exclusiones y duplicados.

**Duplicados de consulta:**

- aplica a nombres que contengan Consulta, Interconsulta o Cuidado/Cuidados;
- agrupa por identificación y fecha sin hora;
- compara únicamente SALDO;
- conserva una fila con el mayor SALDO (en empate, `max()` conserva la primera encontrada);
- elimina las demás de PRE y PEDIR FAC y las coloca en DUPLICADOS;
- DUPLICADOS conserva columnas base más `RAZON DUPLICADO` = `CONSULTA DUPLICADA MISMO PACIENTE Y FECHA`.

**Cálculo:**

- consulta, interconsulta, cuidado, junta médica o equipo interdisciplinario no se calculan como procedimiento y muestran observación de consulta/cuidados;
- para procedimiento COOSALUD, busca rol/porcentaje HOSVI;
- selecciona tarifa CIRUJANO si el rol contiene `CIR`, o AYUDANTE 2 si contiene `AYUD` y `2`;
- distribuye en 100%, 75%, 70%, 60% o 50%;
- `TOTAL COOSALUD` es la suma de los buckets;
- valor clínica = total × porcentaje configurable (70% predeterminado);
- diferencia = SALDO − valor clínica;
- actualiza `VLR. A AUTORIZAR` con valor clínica.

## 7.3 Urología

**Archivo:** `specialties/urologia.py`  
**Entrada:** hoja principal, COOSALUD, HOSVIREPORT; PORCENTAJE_PAGO opcional.  
**Porcentaje base:** 90% clínica.

**Anulados:** agrupa por paciente + código + procedimiento + fecha exacta convertida a texto; ignora factura desde la única pasada. Si el saldo neto tiene valor absoluto menor a 0,01, todo el grupo va a ANULADOS.

**Duplicados:**

- se detectan sobre el archivo original antes de exclusiones;
- solo nombres que coincidan con Consulta, Interconsulta, Junta Médica o Equipo Interdisciplinario;
- agrupa por nombre del paciente + procedimiento exacto + fecha sin hora;
- conserva la primera fila y marca las posteriores, sin comparar SALDO;
- razón: `Consulta/interconsulta duplicada en la misma fecha del servicio`;
- Cuidado no entra en esta regla de duplicados;
- ANULADOS conserva prioridad porque el procesador lo evalúa antes.

**Cálculo:** igual base HOSVI/COOSALUD que Ortopedia, con diferencias:

- 1.75 distribuye 100% + 75%;
- 1.05 distribuye 100% + 50%;
- porcentajes estándar 1.00, 0.75, 0.70, 0.60 y 0.50;
- valor clínica = total × 90% predeterminado.

## 7.4 Cardiología

**Archivo:** `specialties/cardiologia.py`  
**Entrada:** hoja principal y COOSALUD; no usa HOSVIREPORT. PORCENTAJE_PAGO puede leerse globalmente pero no interviene en su cálculo.  
**Porcentajes base:** 90% consultas/cuidados y 36% procedimientos.

**Anulados:** misma estructura de dos pasadas que Ortopedia.

**Duplicados:** misma estructura sincronizada que Ortopedia: paciente + fecha sin hora, Consulta/Interconsulta/Cuidado, mayor SALDO, exclusión simultánea de PRE y PEDIR FAC.

**Cálculo:**

- si el contrato aplica, usa siempre tarifa CIRUJANO de COOSALUD como base;
- Consulta, Interconsulta, Cuidado/Cuidados y la frase “Participación en Junta Médica o Equipo Interdisciplinario” reciben porcentaje de consultas;
- los demás procedimientos reciben porcentaje de procedimientos;
- total = componente consulta + componente procedimiento;
- diferencia = SALDO − total;
- actualiza `VLR. A AUTORIZAR` con el total;
- si no hay tarifa, observación `Sin tarifa COOSALUD` y valores cero;
- si el contrato no aplica, `NO APLICA CARDIOLOGIA`.

## 7.5 Gastroenterología

**Archivo:** `specialties/gastroenterologia.py`  
**Entrada:** hoja principal, COOSALUD, HOSVIREPORT; PORCENTAJE_PAGO opcional.  
**Porcentaje base:** 70% clínica.

**Anulados:** exige columnas factura, código, procedimiento, saldo y estado. Agrupa por factura + paciente + código + procedimiento + fecha. Solo anula cuando:

- el grupo tiene al menos dos filas;
- saldo neto menor a 0,01;
- existe al menos un estado que contiene `ANUL`;
- existe al menos un estado no anulado (o el conjunto de estados está vacío, aunque la condición previa de `has_anulado` impide anular si todos están vacíos).

**Duplicados:** no implementados; no crea DUPLICADOS.

**Cálculo:** misma distribución estándar que Ortopedia y porcentaje clínica 70%.

## 7.6 Emparejamiento HOSVIREPORT compartido

Para Ortopedia, Urología y Gastroenterología:

1. intenta `(factura, código)`;
2. si no encuentra, intenta `(documento, código, fecha)`;
3. usa cursores para consumir coincidencias sucesivas y no reutilizar siempre el primer registro;
4. ignora filas HOSVI cuyo estado normalizado sea exactamente `ANULADO`;
5. si existe PORCENTAJE_PAGO, reemplaza el porcentaje HOSVI para coincidencias por factura/código;
6. si falta tarifa o rol reconocido, agrega observación;
7. si falta porcentaje, agrega `Sin porcentaje (HOSVI / override)`.

---

# 8. Inventario de validaciones

| Validación | Lugar | Verifica | Falla / acción | Mensaje o destino |
|---|---|---|---|---|
| Entrada seleccionada | GUI `process()` | Ruta no vacía | Detiene | “Archivo de entrada requerido”. |
| Salida seleccionada | GUI `process()` | Ruta no vacía | Detiene | “Archivo de salida requerido”. |
| Entrada para validación | GUI `validate_file()` | Ruta no vacía y existente | Detiene | “Archivo no seleccionado/encontrado”. |
| Porcentaje presente | `_get_percent_value()` | Texto no vacío | Detiene | “Debes ingresar…” |
| Rango de porcentaje | `_get_percent_value()` | >0 y ≤100 | Detiene | “…debe estar entre 1 y 100.” |
| Formato numérico | Conversión `float()` | Número válido | Detiene | Genérico “Parámetros inválidos” con detalle. |
| Hojas preliminares | GUI | COOSALUD; además HOSVIREPORT salvo Cardiología | Advierte | “Validación incompleta”. |
| Libro legible | OpenPyXL | Formato Excel accesible | Detiene | Error técnico en validar/procesar. |
| HOSVIREPORT obligatorio | Procesador | Hoja presente para no Cardiología | Detiene | `No existe la hoja 'HOSVIREPORT'…` |
| COOSALUD obligatorio | Procesador | Hoja presente | Detiene | `No existe la hoja 'COOSALUD'…` |
| Cabecera COOSALUD | `build_tarifario_coosalud()` | Encabezados con CIRUJ y AYUD | Detiene | No pudo detectar encabezados. |
| Columnas tarifa | Misma función | CIRUJANO y AYUDANTE 2 | Detiene | Faltan columnas de tarifa. |
| Cabecera HOSVI | `build_hosvireport()` | FACTURA y COD_PROCED | Detiene | No pudo detectar encabezados. |
| Código HOSVI | Misma función | Columna COD_PROCED | Detiene | No encontró COD_PROCED. |
| Cabecera origen | Procesador | CONTR, FAC, PROCED, COD en primeras 200 filas | Detiene | No pudo detectar encabezados. |
| Columnas obligatorias origen | Procesador | NO FAC, CONTRATO, CÓDIGO PROCEDIMIENTO | Detiene | Faltan columnas obligatorias. |
| Fila resumen vacía | Procesador | Factura, contrato, código y procedimiento vacíos | Omite fila | Sin mensaje. |
| Exclusiones comunes | `SpecialtyBase` | 7 condiciones de negocio | Mueve a AMARILLO | Razón específica. |
| Anulados | Especialidad | Grupo y saldo neto | Mueve a ANULADOS y corta flujo | `Grupo anulable (neto saldo=0)` + estado si existe. |
| Duplicados | Especialidad | Regla particular | Mueve a DUPLICADOS | Razón particular. |
| Sincronía PRE/PEDIR | Ortopedia/Cardiología | Igual número de filas temporales | Detiene | `Las filas PRE y PEDIR FACT no estan sincronizadas`. |
| Archivo último generado | GUI | Resultado en sesión y ruta existente | Informa | Archivo no disponible/no encontrado. |
| Manual | GUI | PDF existe y puede abrirse | Informa error | Manual no encontrado/no fue posible abrir. |
| Carpeta de salida | GUI | Ruta y carpeta existentes | Advierte | Ubicación no disponible/no encontrada. |

**Validación parcial identificada:** el botón “Validar archivo” solo verifica nombres de hojas. No garantiza columnas, tarifas, encabezados ni reglas; esos fallos aparecen al procesar.

---

# 9. Inventario de errores y solución

| Error observado o posible | Posible causa comprobable | Responsable | Acción recomendada |
|---|---|---|---|
| Archivo de entrada requerido | Campo vacío | GUI | Seleccionar Excel. |
| Archivo de salida requerido | Campo vacío | GUI | Definir `.xlsx`. |
| Archivo no encontrado | Ruta movida/eliminada | GUI/SO | Volver a seleccionar. |
| Parámetros inválidos | Vacío, texto no numérico, 0 o >100 | GUI | Corregir porcentaje entre 1 y 100. |
| No existe HOSVIREPORT | Hoja ausente o nombre diferente | Procesador | Agregar/renombrar hoja; no aplica a Cardiología. |
| No existe COOSALUD | Hoja ausente o nombre diferente | Procesador | Agregar/renombrar hoja. |
| No detecta encabezados COOSALUD | No existen textos CIRUJ/AYUD en primeras 200 filas | Excel utils | Ajustar encabezados. |
| Faltan columnas tarifa | Falta CIRUJANO o AYUDANTE 2 reconocible | Excel utils | Corregir columnas. |
| No detecta HOSVIREPORT | Falta FACTURA/COD_PROCED | Excel utils | Corregir cabecera. |
| No detecta encabezados origen | Falta alguno de CONTR/FAC/PROCED/COD o está después de fila 200 | Procesador | Ajustar hoja principal. |
| Faltan columnas obligatorias | NO FAC, CONTRATO o código no reconocido | Procesador | Usar un encabezado admitido. |
| Especialidad no soportada | Nombre CLI fuera del catálogo | Factory | Usar uno de los cuatro nombres exactos. |
| Hoja original inexistente | `--original-sheet` inválido | OpenPyXL/procesador | Corregir nombre; GUI siempre usa primera hoja. |
| Sin tarifa / rol no reconocido | Código ausente o rol HOSVI no contiene CIR/AYUD+2 | Especialidad | Revisar COOSALUD/HOSVIREPORT; la fila no necesariamente detiene el proceso. |
| Sin porcentaje HOSVI/override | Coincidencia sin porcentaje | Especialidad | Completar `%_FACTURADO` o PORCENTAJE_PAGO. |
| Porcentaje no esperado | HOSVI distinto de buckets soportados | Especialidad | Corregir porcentaje según especialidad. |
| PRE y PEDIR desincronizadas | Inconsistencia interna de listas | Ortopedia/Cardiología | Revisión técnica; el sistema detiene para no producir salida inconsistente. |
| Error al guardar | Ruta sin permisos, archivo bloqueado o carpeta inexistente | OpenPyXL/SO | Cerrar Excel, elegir ruta válida y con permisos. |
| No fue posible abrir ubicación/archivo/manual | Asociación del SO, ruta o recurso inválido | GUI/SO | Verificar existencia y aplicaciones predeterminadas. |
| Configuración no persistida | AppData no escribible o JSON corrupto | `settings.py` | Revisar `%LOCALAPPDATA%/HONORARIOS`; un JSON ilegible vuelve a defaults en memoria. |
| Conteos visuales mayores que datos | `ProcessResult` usa `max_row - 1` después de agregar cuadros resumen | Procesador/escritor | Limitación actual: interpretar métricas con cautela; requiere corrección futura, no está corregido en esta fase. |
| Validación dice compatible pero proceso falla | Botón solo revisa hojas | GUI | Considerar la validación como preliminar; revisar encabezados y contenido. |
| XLSM/XLTX aceptado pero macros/plantilla no preservadas | Diálogo acepta extensiones, carga sin `keep_vba` y salida es XLSX | GUI/OpenPyXL | Trabajar sobre copia y asumir salida sin macros. |

La GUI captura cualquier excepción del procesamiento y la presenta como “Procesamiento interrumpido” con detalle técnico. La CLI no encapsula estas excepciones.

---

# 10. Archivos de entrada

## 10.1 Libro principal

- **Formato admitido por diálogo:** XLSX, XLSM, XLTX, XLTM.
- **Formato procesado:** cualquiera que OpenPyXL pueda abrir; salida siempre XLSX.
- **Nombre:** libre.
- **Obligatorio:** sí.
- **Hoja base:** primera hoja en GUI; configurable por CLI.
- **Cabecera:** debe estar dentro de las primeras 200 filas.
- **Columnas mínimas formalmente obligatorias:** NO FAC, CONTRATO y CÓDIGO PROCEDIMIENTO.
- **Columnas funcionalmente necesarias:** PROCEDIMIENTO, SALDO, identificación, nombre, fecha, estado, tipo de factura y VLR. A AUTORIZAR afectan reglas/cálculos. Si faltan algunas, el proceso puede degradarse, excluir filas o no crear fórmulas en vez de fallar de inmediato.

Alias reconocidos principales:

- factura: NO FAC, NUM FAC, N° FAC, FACTURA;
- código: CÓDIGO PROCEDIMIENTO, COD PROCEDIMIENTO, CÓDIGO PROC, COD_PROCED, COD;
- paciente: varias formas de IDENTIFICACIÓN/ID/DOC/DOCUMENTO;
- fecha: FECHA DEL SERVICIO, FECHA SERVICIO, FECHA_PROCED, FECHA PROCED, FECHA;
- nombre: NOMBRE PACIENTE, NOMBRE_PACIENTE, NOMBRE P., PACIENTE.

## 10.2 Hoja COOSALUD

- **Obligatoria:** todas las especialidades.
- **Nombre predeterminado:** `COOSALUD`.
- **Cabecera:** contiene CIRUJ y AYUD en primeras 200 filas.
- **Código:** CUPS/CÓDIGO/COD o primera columna como fallback.
- **Valores:** tarifa CIRUJANO y AYUDANTE 2.

## 10.3 Hoja HOSVIREPORT

- **Obligatoria:** Ortopedia, Urología y Gastroenterología.
- **No requerida:** Cardiología.
- **Nombre:** `HOSVIREPORT`.
- **Mínimo estructural:** FACTURA y COD_PROCED.
- **Campos usados si existen:** rol/honorario, porcentaje facturado, documento, fecha y estado.

## 10.4 Hoja opcional PORCENTAJE_PAGO

- **Obligatoria:** no.
- **Propósito:** sobrescribir porcentaje HOSVI por factura + código.
- **Cabecera esperada:** referencias a FACT, COD y PORC.
- **Campos:** factura, código y porcentaje.
- **Aplicación:** coincidencia principal por factura/código; no se aplica explícitamente al índice alternativo documento/código/fecha.

## 10.5 Configuración persistente

No es un archivo que el usuario cargue. El sistema crea `%LOCALAPPDATA%/HONORARIOS/settings.json` con:

- Ortopedia 70%;
- Urología 90%;
- Cardiología 90% consultas y 36% procedimientos;
- Gastroenterología 70%.

---

# 11. Archivos de salida

## 11.1 Libro XLSX

Nombre automático: `PEDIR FAC_<ESPECIALIDAD>_<NOMBRE_ENTRADA>.xlsx`.

## 11.2 Hojas comunes

| Hoja | Contenido | Columnas |
|---|---|---|
| PRE | Filas válidas procesadas | Columnas originales + columnas calculadas de especialidad. |
| AMARILLO | Filas excluidas | Columnas originales + RAZON EXCLUSION. |
| PEDIR FAC | Mismas filas válidas finales que PRE para el flujo sincronizado de Ortopedia/Cardiología; base administrativa | Solo columnas originales, con VLR. A AUTORIZAR actualizado cuando corresponde. |
| ANULADOS | Grupos anulados | Columnas originales + RAZON ANULADO. |

## 11.3 Hoja DUPLICADOS

- Ortopedia: sí.
- Cardiología: sí.
- Urología: sí, con regla y razón diferentes.
- Gastroenterología: no.

Contiene columnas originales más `RAZON DUPLICADO`; no incluye columnas calculadas.

## 11.4 Columnas calculadas PRE

**Ortopedia/Urología/Gastroenterología:** ROL HOSVI, % HOSVI, tarifas CIRUJANO/AYUDANTE 2, buckets 100/75/70/60/50, TOTAL COOSALUD, `<porcentaje>% CLINICA`, DIF SALDO-COOSALUD, OBS.

**Cardiología:** 100% COOSALUD FORMULADO, `% DE COOSALUD` de consultas, `% PROCED.COOSALUD`, TOTAL COOSALUD, DIFE. HOSV Y COOSALUD, OBS.

## 11.5 Resúmenes y formato

- PRE agrega TOTAL BAJAR HOSV, TOTAL COOSALUD y DIF con fórmulas y borde rojo.
- PEDIR FAC agrega TOTAL PEDIR FACT.
- encabezados azul claro/negrita;
- panel congelado A2;
- ancho automático.

**Hallazgo QA:** los conteos devueltos se calculan después de agregar las filas de resumen. Por tanto, cuando hay datos, `pre_rows` y `pedir_rows` no representan estrictamente el número de registros; incluyen la posición ocupada por resúmenes y espacios. Las hojas conservan los datos, pero las métricas y mensajes pueden estar inflados.

---

# 12. Capturas de pantalla requeridas para la Fase 2

| N.º | Captura | Motivo | Capítulo futuro |
|---|---|---|---|
| 1 | Ventana principal completa, estado inicial | Identificar áreas | Conociendo la interfaz |
| 2 | Vista en ventana estrecha/apilada | Mostrar diseño adaptable y scroll | Requisitos e interfaz |
| 3 | Selector de especialidad desplegado | Mostrar catálogo exacto | Selección de especialidad |
| 4 | Modo porcentajes simple | Ortopedia/Urología/Gastro | Configuración de porcentajes |
| 5 | Modo Cardiología | Dos porcentajes independientes | Configuración de Cardiología |
| 6 | Diálogo Buscar archivo | Extensiones y selección | Preparación del archivo |
| 7 | Ruta de salida sugerida | Convención automática | Archivo de salida |
| 8 | Diálogo Guardar como | Salida manual | Archivo de salida |
| 9 | Validación exitosa | Confirmación preliminar | Validar archivo |
| 10 | Validación con hojas faltantes | Recuperación operativa | Solución de problemas |
| 11 | Advertencia de porcentaje inválido | Rango permitido | Validaciones |
| 12 | Procesamiento en curso | Controles deshabilitados, progreso y log | Ejecutar procesamiento |
| 13 | Bitácora con etapas | Interpretar seguimiento | Bitácora |
| 14 | Procesamiento completado | Resumen final | Finalización |
| 15 | Confirmación para guardar porcentaje simple | Persistencia | Configuración |
| 16 | Confirmación cardiológica | Persistencia dual | Configuración Cardiología |
| 17 | Resumen del último proceso | Lectura y contenido | Herramientas |
| 18 | Copiar resumen confirmado | Uso del portapapeles | Herramientas |
| 19 | Ayuda de uso | Guía rápida integrada | Soporte |
| 20 | Guía del módulo | Alcance integrado | Soporte |
| 21 | Acerca del sistema | Versión/tecnologías | Información del sistema |
| 22 | Manual abierto en visor externo | Acceso al recurso | Soporte |
| 23 | Explorador seleccionando archivo generado | Ubicación de salida | Resultados |
| 24 | Libro final con pestañas | Mapa de hojas | Interpretación de resultados |
| 25 | PRE Ortopedia/Urología/Gastro | Columnas HOSVI/COOSALUD | Hoja PRE |
| 26 | PRE Cardiología | Columnas 90/36 | Hoja PRE Cardiología |
| 27 | Cuadro resumen PRE | Fórmulas finales | Totales |
| 28 | PEDIR FAC y total | Gestión administrativa | Hoja PEDIR FAC |
| 29 | AMARILLO con razones | Exclusiones | Hoja AMARILLO |
| 30 | ANULADOS con razón/estado | Anulaciones | Hoja ANULADOS |
| 31 | DUPLICADOS Ortopedia/Cardiología | Regla mayor SALDO | Duplicados |
| 32 | DUPLICADOS Urología | Regla primera ocurrencia | Duplicados Urología |
| 33 | Error de procesamiento representativo | Interpretar detalle técnico | Solución de problemas |

El PDF histórico ya contiene algunas capturas de una interfaz y libros anteriores, pero deben regenerarse contra la versión actual para evitar documentar estados obsoletos.

---

# 13. Índice detallado del futuro Manual de Usuario

1. Presentación
   1.1 Propósito del manual
   1.2 Alcance del sistema
   1.3 Usuarios destinatarios
   1.4 Convenciones utilizadas
2. Descripción del aplicativo
   2.1 Objetivo del módulo Facturados
   2.2 Especialidades soportadas
   2.3 Entradas, procesamiento y resultados
   2.4 Limitaciones conocidas de la versión
3. Requisitos
   3.1 Sistema operativo y ejecutable
   3.2 Software para visualizar Excel y PDF
   3.3 Permisos de lectura/escritura
   3.4 Conocimientos operativos recomendados
4. Preparación del archivo de entrada
   4.1 Formatos admitidos
   4.2 Hoja principal
   4.3 Encabezados y columnas reconocidas
   4.4 Hoja COOSALUD
   4.5 Hoja HOSVIREPORT
   4.6 Hoja opcional PORCENTAJE_PAGO
   4.7 Requisitos por especialidad
   4.8 Lista de comprobación previa
5. Inicio del sistema
   5.1 Ejecución del programa
   5.2 Carga de configuración
   5.3 Estado inicial
6. Conociendo la interfaz
   6.1 Encabezado
   6.2 Panel de métricas
   6.3 Parámetros de procesamiento
   6.4 Soporte y consulta
   6.5 Herramientas
   6.6 Barra de progreso
   6.7 Bitácora
   6.8 Diseño adaptable y desplazamiento
7. Configuración del proceso
   7.1 Seleccionar especialidad
   7.2 Porcentaje base y porcentaje temporal
   7.3 Porcentajes de Cardiología
   7.4 Restablecer porcentajes
   7.5 Guardar nuevos valores base
8. Selección de archivos
   8.1 Buscar archivo de entrada
   8.2 Nombre automático de salida
   8.3 Definir salida manual
   8.4 Extensiones y conservación de macros
9. Validación previa
   9.1 Ejecutar Validar archivo
   9.2 Hojas requeridas por especialidad
   9.3 Interpretar resultado correcto
   9.4 Corregir hojas faltantes
   9.5 Alcance real de la validación preliminar
10. Ejecución del procesamiento
    10.1 Iniciar
    10.2 Estados y porcentajes de progreso
    10.3 Controles durante la ejecución
    10.4 Interpretar la bitácora
    10.5 Finalización correcta
    10.6 Persistencia de porcentajes
11. Reglas generales
    11.1 Prioridad de ANULADOS
    11.2 Exclusiones y AMARILLO
    11.3 Contratos COOSALUD aplicables
    11.4 Tarifas y roles HOSVI
    11.5 Override PORCENTAJE_PAGO
12. Reglas de Ortopedia
    12.1 Anulados por factura
    12.2 Anulados entre facturas
    12.3 Consultas duplicadas y mayor SALDO
    12.4 Cálculo del procedimiento
    12.5 Observaciones
13. Reglas de Urología
    13.1 Anulados/refacturados
    13.2 Duplicados de consulta
    13.3 Distribuciones estándar
    13.4 Casos 175% y 105%
    13.5 Valor clínica
14. Reglas de Cardiología
    14.1 Anulados por factura
    14.2 Anulados entre facturas
    14.3 Duplicados sincronizados
    14.4 Consultas/cuidados/junta médica
    14.5 Procedimientos al 36%
    14.6 Valores configurables
15. Reglas de Gastroenterología
    15.1 Anulados con estado
    15.2 Cálculo HOSVI/COOSALUD
    15.3 Consultas y cuidados sin cálculo
16. Archivo de salida
    16.1 Convención de nombre
    16.2 Orden y estructura de hojas
    16.3 PRE
    16.4 PEDIR FAC
    16.5 AMARILLO
    16.6 ANULADOS
    16.7 DUPLICADOS
    16.8 Totales y fórmulas
    16.9 Formato, congelación y anchos
17. Herramientas posteriores
    17.1 Abrir ubicación
    17.2 Abrir último generado
    17.3 Ver resumen
    17.4 Copiar resumen
    17.5 Consultar bitácora
18. Soporte integrado
    18.1 Ayuda de uso
    18.2 Guía del módulo
    18.3 Manual del módulo
    18.4 Acerca del sistema
19. Mensajes y validaciones
    19.1 Informativos
    19.2 Advertencias
    19.3 Confirmaciones
    19.4 Errores técnicos
20. Solución de problemas
    20.1 Archivo no abre
    20.2 Hojas faltantes
    20.3 Encabezados no detectados
    20.4 Tarifas o porcentajes faltantes
    20.5 Archivo bloqueado o sin permisos
    20.6 Resultados inesperados
    20.7 Manual/archivo no disponible
    20.8 Conteos visuales y filas de resumen
21. Buenas prácticas operativas
    21.1 Conservar copia del original
    21.2 Validar especialidad
    21.3 Revisar porcentajes
    21.4 Revisar AMARILLO, ANULADOS y DUPLICADOS
    21.5 Verificar totales
22. Anexos
    22.1 Diccionario de hojas
    22.2 Diccionario de columnas
    22.3 Matriz comparativa de especialidades
    22.4 Catálogo de mensajes
    22.5 Lista de capturas

---

# 14. Hallazgos transversales para QA, arquitectura y UX

1. **Arquitectura:** separación clara entre UI, orquestación, utilidades y reglas; las especialidades extienden `SpecialtyBase` y el procesador usa capacidades opcionales mediante `hasattr`.
2. **Duplicación técnica:** Ortopedia y Cardiología contienen implementaciones prácticamente idénticas para anulados y duplicados; no se centralizan en Base por decisión actual.
3. **Módulos incompletos:** `validators.py`, `module_factory.py` y `config_service.py` están vacíos.
4. **Validación superficial:** “Validar archivo” no valida columnas ni contenido.
5. **Conteos defectuosos:** las métricas se basan en `max_row` después de resúmenes.
6. **Configuración:** el JSON del repositorio no es consumido; la fuente efectiva es el JSON de AppData creado desde constantes Python.
7. **Interfaz bloqueante:** el proceso corre en el hilo de la UI y llama `root.update()`; deshabilita controles, pero no usa hilo de trabajo ni cancelación.
8. **Bitácora editable:** el usuario puede modificar el texto visual de log, aunque esas ediciones no afectan el procesamiento.
9. **Sin cancelación:** no existe botón para detener un proceso iniciado.
10. **Sin historial persistente:** `last_result` solo vive durante la sesión.
11. **Macros:** la selección admite XLSM, pero no se preserva VBA y la salida siempre es XLSX.
12. **Manual incluido:** es funcional como acceso, pero su contenido no refleja plenamente las reglas actuales y debe reemplazarse en Fase 2.
13. **Pruebas parciales:** hay pruebas de Cardiología y Ortopedia para las mejoras recientes; faltan pruebas automatizadas de UI, Urología, Gastroenterología, configuración y errores de estructura.
14. **CLI parcialmente documentada:** existe y permite elegir hoja original/nombres HOSVI/COOSALUD, pero no se expone en la GUI ni en el manual histórico.
15. **Recurso icono:** está integrado en el ejecutable, no en la ventana mediante una llamada Tkinter explícita.

Este levantamiento es el insumo de Fase 1. No contiene instrucciones narrativas de usuario final ni desarrolla todavía los capítulos del manual.
