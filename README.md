# ⚽ Generador de Fixture Multicategoría

Aplicación de escritorio que genera automáticamente fixtures deportivos bajo modalidad **todos contra todos**, optimizando la coincidencia de rivales entre múltiples categorías.

Diseñado para clubes que manejan varias divisiones (ej: Primera, Reserva, Juveniles, etc.) y necesitan organizar fechas de forma eficiente y sin errores manuales.

---

## 🚀 ¿Qué hace este programa?

* Genera fixtures completos automáticamente
* Soporta **múltiples categorías dinámicas** (no están hardcodeadas)
* Minimiza fechas libres
* Optimiza para que un club enfrente al **mismo rival en todas sus categorías** en una misma fecha
* Exporta resultados en Excel con:

  * Fixtures por categoría
  * Calendario unificado
  * Partidos por fecha (formato imprimible)
  * Resumen estadístico

---

## 🧠 Lógica del sistema

El motor utiliza **optimización combinatoria** con Google OR-Tools:

* Round-robin por categoría
* Restricción: cada equipo juega una vez por fecha
* Restricción: cada par de equipos se enfrenta una sola vez
* Función objetivo: maximizar coincidencia de rivales entre categorías

---

## 🖥️ Interfaz

Aplicación simple, pensada para usuarios no técnicos:

* Selección de Excel base
* Generación automática de fixture
* Apertura directa del resultado

![Interfaz del programa](assets/app.png)

---

## 📂 Estructura del proyecto

```
fixture-club/
│
├── app.py                  # Interfaz gráfica (Tkinter)
├── main.py                 # Motor de generación (OR-Tools)
├── Plantilla_Clubes_REAL.xlsx
│
├── dist/
│   └── app/                # Ejecutable listo para usar (.exe)
│
└── assets/
    └── app.png             # Captura de la interfaz
```

---

## ⚙️ Tecnologías utilizadas

* Python 3
* OR-Tools (optimización)
* Pandas (procesamiento de datos)
* OpenPyXL (exportación Excel)
* Tkinter (interfaz gráfica)
* PyInstaller (empaquetado a .exe)

---

## 📊 Formato de entrada

El programa utiliza un Excel como base de datos:

| EQUIPO     | S | SS | M |
| ---------- | - | -- | - |
| San Martín | X | X  | X |
| Campito    | X | X  |   |
| ...        |   | X  | X |

* Cada columna representa una categoría
* Las categorías son detectadas automáticamente
* Se marca con **X** si el club participa

---

## ▶️ Cómo usar

### Opción 1 (recomendada)

Ejecutar el programa:

```
dist/app/app.exe
```

### Opción 2 (modo desarrollo)

```
python app.py
```

---

## 📦 Instalación (desarrollo)

```bash
pip install pandas openpyxl ortools
```

---

## 💡 Características destacadas

* Adaptable a cualquier cantidad de categorías
* No requiere conocimientos técnicos
* Evita errores humanos en la planificación
* Genera salidas listas para imprimir o compartir

---

## 🧩 Futuras mejoras

* Exportación a PDF por fecha
* Historial de fixtures
* Configuración avanzada de restricciones
* Interfaz con tema visual personalizable

---

Desarrollado para resolver un problema real en la organización de torneos deportivos.
