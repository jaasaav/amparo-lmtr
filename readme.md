# Generador de Demanda de Amparo Indirecto – LMTR y Lineamientos (v1.0) (16/ago/26)

## Mejoras implementadas en esta versión

1. **Control horario y ventana temporal precisa**
* Incorporación de selector de hora (`st.time_input`) para registrar el momento exacto de la notificación o suspensión.


* Cálculo dinámico de la fecha base según el último dígito del número telefónico y la ventana transitoria de 72 horas.




2. **Formato tipográfico estricto y protección de notas al pie**
* Función `aplicar_formato_proemio_seguro_con_notas` que aplica negritas selectivas en el proemio (personería, domicilio, código postal, ciudad, usuario del portal y datos de abogados) **sin destruir los nodos XML de las notas al pie de página** (como la cita de la nota 1 en el apartado de derecho).
* Función `renumerar_y_formatear_pruebas` que renumera secuencialmente el capítulo probatorio, resalta los títulos de las documentales y destaca elementos clave como la concesionaria, el número telefónico, el banco y el nivel de cuenta.


3. **Sanitización de nombres de archivo para el Portal del PJF**
* Implementación de la función `limpiar_nombre_archivo` basada en normalización Unicode y expresiones regulares, eliminando acentos, espacios, guiones y caracteres especiales para generar nombres de archivo estrictamente alfanuméricos que evitan rechazos en el sistema del Poder Judicial de la Federación.


4. **Estructura modular y comunidad de código abierto**
* Interfaz de 7 pestañas optimizada para la captura de datos del quejoso, representación de menores de edad (art. 8 de la Ley de Amparo), selección de concesionaria, cuentas bancarias y suplencia de la queja.
* Integración de guías detalladas para el trámite de la **FIREL** y presentación en línea, así como una sección abierta de convocatoria para colaboración mediante *Pull Requests*.



## Cómputo automático del plazo

15 días hábiles (arts. 17 y 18 LAmp + 35/38 LFPA):

* Excluye sábados y domingos.


* Días inhábiles judiciales 2026: 14-16 sep, 12 oct, 2/16/20 nov, 16-31 dic.



## Cómo ejecutar

```bash
pip install -r requirements.txt
streamlit run app.py --server.port 8501

```

Abra `http://localhost:8501` en su navegador web.

## Archivos del proyecto

| Archivo | Descripción |
| --- | --- |
| `app.py` | Aplicación principal de Streamlit (v1.0) |
| `amparo_template_clean.docx` | Plantilla Word optimizada y limpia de metadatos o comentarios |
| `requirements.txt` | Dependencias del proyecto (`streamlit`, `python-docx`) |
| `LICENCIA.md` | Términos de la licencia CC BY-SA 4.0 y excepciones profesionales |
| `CONTRIBUTING.md` | Guía para la colaboración abierta y envío de mejoras |

## Competencia

Juzgado de Distrito en materia de Acceso a la Información Pública y Protección de Datos Personales con residencia en Aguascalientes, Ags. (Acuerdo General 8/2025 y Contradicción de Tesis 37/2019 de la SCJN).

## Créditos y autoría

* **Dirección jurídica y supervisión:** Julio Amador.


* **Desarrollo técnico e IA colaborativa:** Asistencia de Gemini, Gemini Notebook y Grok.



## Aviso legal

Herramienta de apoyo técnico e informático. Se recomienda la revisión y validación final por un abogado postulante autorizado antes de la presentación formal del escrito ante los órganos jurisdiccionales federales.