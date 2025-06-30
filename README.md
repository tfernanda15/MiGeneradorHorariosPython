# Generador de Horarios de Clases con Python

Este proyecto es un script en Python diseñado para generar horarios de clases para profesores y aulas, gestionando la disponibilidad y evitando conflictos. El resultado es un horario estructurado exportado a un archivo Excel.

## Características

* **Generación de Horarios:** Asigna clases a franjas horarias, profesores y aulas disponibles.
* **Manejo de Conflictos:** Evita asignaciones múltiples para un mismo profesor o aula en el mismo bloque horario.
* **Validación de Capacidad:** Considera la capacidad de las aulas para el número de estudiantes de cada grupo.
* **Exportación a Excel:** Los horarios generados se exportan a un archivo `.xlsx` legible y organizado.

## Tecnologías Utilizadas

* Python 3.x
* `openpyxl` (para la manipulación de archivos Excel)

## Cómo Empezar / Cómo Ejecutar el Proyecto

Sigue estos pasos para poner en marcha el generador de horarios en tu máquina local:

1.  **Clonar el Repositorio:**
    ```bash
    git clone [https://github.com/tfernanda15/MiGeneradorHorariosPython.git](https://github.com/tfernanda15/MiGeneradorHorariosPython.git)
    cd MiGeneradorHorariosPython
    ```

2.  **Crear y Activar un Entorno Virtual (Recomendado):**
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate  # En Linux/macOS
    # .venv\Scripts\activate   # En Windows (CMD)
    # .venv\Scripts\Activate.ps1 # En Windows (PowerShell)
    ```

3.  **Instalar Dependencias:**
    ```bash
    pip install openpyxl
    ```

4.  **Ejecutar el Script:**
    ```bash
    python generador_horarios.py
    ```
    El script imprimirá el progreso en la consola y generará un archivo `horario_escolar.xlsx` en el mismo directorio.

## Estructura de los Datos (Para modificar y adaptar)

Puedes encontrar las estructuras de datos (profesores, aulas, materias, clases requeridas) en el código fuente `generador_horarios.py`. Modifícalas directamente en el archivo para adaptar el generador a tus necesidades específicas.

## Ejemplo de Salida

Un ejemplo del horario generado puede ser visualizado abriendo el archivo `horario_escolar.xlsx` que se incluye en este repositorio.

## Autor

* [Tu Nombre Completo o Tu Usuario de GitHub: tfernanda15]

## Posibles Mejoras Futuras

* Creación de una Interfaz Gráfica de Usuario (GUI) o una aplicación web para facilitar la interacción.
* Implementación de lógica para cargar datos desde archivos externos (CSV, JSON).
* Integración de restricciones más avanzadas (ej. disponibilidad horaria específica de profesores, asignación de aulas preferidas).
