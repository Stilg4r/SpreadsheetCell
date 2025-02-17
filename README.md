
# Clase Auxiliar para ExcelJS

Esta librería proporciona una serie de utilidades para manipular hojas de cálculo utilizando ExcelJS. Entre las funcionalidades principales se encuentran:

- **Conversión de columnas:**  
  - `columnToNumber`: Convierte una columna en formato alfabético (por ejemplo, "A" o "AB") a su equivalente numérico.  
  - `numberToColumn`: Realiza la conversión inversa, pasando de un número a la notación de columnas de Excel.

- **Manipulación de secuencias de columnas:**  
  - `incrementColumn`: Permite incrementar una columna en una cantidad determinada.  
  - `generateColumnSequence`: Función generadora que produce la secuencia de columnas de forma infinita.

- **Aplicación de estilos de borde a celdas:**  
  - `borderBox`: Aplica estilos de borde diferenciados para los bordes externos e internos de un rango de celdas dentro de una hoja de cálculo.

- **Clase SpreadsheetCell:**  
  La clase `SpreadsheetCell` es el núcleo de esta librería. Proporciona métodos para:
  - Crear y navegar entre celdas de una hoja de cálculo.
  - Aplicar y gestionar formatos de celdas.
  - Manipular la posición de la celda actual (mover a la siguiente/previa columna o fila).
  - Fusionar celdas y rellenar datos en filas, columnas o tablas completas.
  - Clonar y destruir instancias para liberar recursos.
  
Esta clase se integra de manera eficiente con ExcelJS, facilitando la manipulación avanzada de hojas de cálculo y permitiendo una gestión modular y extensible de las celdas.


## Uso Básico

```javascript
import { SpreadsheetCell, borderBox } from './ruta-a-su-clase-auxiliar';

// Suponga que 'worksheet' es una instancia válida de ExcelJS
const cellHelper = new SpreadsheetCell({ column: "A", row: 1, worksheet });

// Establecer un valor y aplicar formato a la celda actual
cellHelper.setAndApplyFormat("Valor de ejemplo", { font: { bold: true } });

// Moverse a la siguiente columna y establecer otro valor
cellHelper.setAndMoveToNextColumn("Otro valor", { font: { italic: true } });

// Aplicar un borde a un rango de celdas
borderBox(worksheet, { row: 1, column: "A" }, { row: 3, column: "C" });
```

## Ejemplos de Uso

### Ejemplo 1: Conversión de Columnas
```javascript
import { columnToNumber, numberToColumn, incrementColumn } from './ruta-a-su-clase-auxiliar';

console.log(columnToNumber("A"));    // Salida: 1
console.log(columnToNumber("AB"));   // Salida: 28
console.log(numberToColumn(28));     // Salida: "AB"
console.log(incrementColumn("A", 5));  // Salida: "F"
```

### Ejemplo 2: Generador de Secuencia de Columnas
```javascript
import { generateColumnSequence } from './ruta-a-su-clase-auxiliar';

const columnGen = generateColumnSequence("X");
console.log(columnGen.next().value); // Salida: "X"
console.log(columnGen.next().value); // Salida: "Y"
console.log(columnGen.next().value); // Salida: "Z"
console.log(columnGen.next().value); // Salida: "AA"
```

### Ejemplo 3: Uso de SpreadsheetCell con ExcelJS
```javascript
import { SpreadsheetCell, borderBox } from './ruta-a-su-clase-auxiliar';
import ExcelJS from 'exceljs';

// Crear un workbook y una hoja de cálculo de ExcelJS
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Ejemplo');

// Crear una instancia de SpreadsheetCell
const cellHelper = new SpreadsheetCell({ column: "A", row: 1, worksheet });

// Establecer un valor y aplicar formato a la celda actual
cellHelper.setAndApplyFormat("Encabezado", { font: { bold: true, size: 14 } });

// Moverse a la siguiente columna y agregar otro valor
cellHelper.setAndMoveToNextColumn("Subencabezado", { font: { italic: true } });

// Rellenar una fila completa con datos
cellHelper.fillRow({ values: ["Dato1", "Dato2", "Dato3"], format: { border: { style: 'thin' } } });

// Moverse a la siguiente fila
cellHelper.moveToNextRow();

// Fusionar celdas (por ejemplo, fusionar 2 columnas a partir de la celda actual)
cellHelper.mergeCells({ columns: 2, rows: 0 });
cellHelper.value = "Celda fusionada";

// Aplicar bordes a un rango de celdas
borderBox(worksheet, { row: 1, column: "A" }, { row: 3, column: "C" });
```

## Documentación

### Funciones de Conversión
- **columnToNumber(column):**  
  Convierte una columna en formato alfabético a su valor numérico equivalente.  
  Ejemplo: `"AB"` se convierte a \(26^1 \times 1 + 26^0 \times 2\).

- **numberToColumn(number):**  
  Convierte un número a la notación alfabética de columnas de Excel.

### Funciones de Navegación
- **incrementColumn(column, number):**  
  Incrementa una columna en la cantidad especificada.

- **generateColumnSequence(column):**  
  Genera una secuencia infinita de columnas a partir de la columna dada.

### Clase SpreadsheetCell
- **Constructor:**  
  Recibe un objeto con propiedades `column`, `row` y `worksheet`.  
  Valida que el formato de la columna y el valor de la fila sean correctos.

- **Métodos de Navegación y Formateo:**  
  Permite mover la celda actual, aplicar formatos, fusionar celdas, rellenar datos y obtener la dirección de la celda.

- **Gestión de Formato:**  
  Métodos como `applyFormatting`, `overwriteFormatting`, `addFormatting` permiten la personalización visual de las celdas.

- **Otras Utilidades:**  
  Métodos para clonar la instancia y destruir referencias internas para mejorar la gestión de memoria.

## Notas

- **Integración con ExcelJS:**  
  Esta clase auxiliar está diseñada para trabajar en conjunto con [ExcelJS](https://www.npmjs.com/package/exceljs) y facilita la manipulación de celdas, permitiendo operaciones complejas de manera sencilla.

- **Generado por IA:**  
  **Nota:** Este README ha sido generado por inteligencia artificial para facilitar su comprensión y uso.
