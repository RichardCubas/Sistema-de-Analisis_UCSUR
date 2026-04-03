📊 Informe Final de Curso – Generador de Reportes Académicos

Sistema desarrollado para la generación automatizada de reportes académicos en Excel y PDF, utilizado en los cursos básicos de la Universidad Científica del Sur.

Permite analizar el rendimiento de los estudiantes a nivel de evaluación, carrera, docente y curso, aplicando reglas académicas reales y criterios institucionales.

🚀 Funcionalidades principales
📈 Cálculo de situación final del estudiante

Fórmula oficial:

FINAL = 0.18·EC1 + 0.20·EP + 0.18·EC2 + 0.19·EC3 + 0.25·EF
Redondeo por evaluación (half-up)
Clasificación automática:
Aprobado (≥ 12.5)
Desaprobado
No rindió
📊 Reportes en Excel (4 hojas)
Resumen por evaluación
Análisis por carrera (situación final)
Análisis por docente
Resumen global del curso
✔ Incluye gráficos automáticos
📄 Reportes en PDF
Tablas dinámicas con ajuste de texto
Gráficos integrados
Formato institucional
Listo para presentación
🧠 Reglas académicas implementadas
Porcentajes:
Aprobados / desaprobados → sobre quienes rindieron
No rindieron → sobre total
Promedios calculados solo con estudiantes que rindieron
Identificación automática de estudiantes sin participación
🧩 Estructura de datos requerida

Archivo en formato Parquet con las columnas:

CodigoAlumno
Alumno
Curso
Seccion
Carrera
Docente
Evaluacion
Nota
🛠️ Tecnologías utilizadas
Python
Pandas / NumPy
Matplotlib
XlsxWriter
FPDF
🎯 Uso

Este sistema es utilizado para:

Seguimiento académico en cursos básicos
Evaluación del desempeño docente
Análisis por carrera
Generación de informes institucionales
