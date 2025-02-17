export const columnToNumber = (column) => {
  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let result = 0;
  for (let i = 0, j = column.length - 1; i < column.length; i++, j--) {
    result += Math.pow(letters.length, j) * (letters.indexOf(column[i]) + 1);
  }
  return result;
};

export const numberToColumn = (number) => {
  let columnLetter = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    number = Math.floor((number - remainder) / 26);
  }
  return columnLetter;
};

export const incrementColumn = (column, number) => {
  return numberToColumn(columnToNumber(column) + number);
};

export const generateColumnSequence = function* (column) {
  let start = columnToNumber(column);
  while (true) {
    yield numberToColumn(start++);
  }
};

export const borderBox = (
  worksheet,
  startCell = { row: 1, column: 1 },
  endCell = { row: 3, column: 3 },
  borderStyle = 'thin',
  innerBorderStyle = 'dotted'
) => {
  const startColumn = columnToNumber(startCell.column);
  const endColumn = columnToNumber(endCell.column);

  for (let row = startCell.row; row <= endCell.row; row++) {
    for (let col = startColumn; col <= endColumn; col++) {
      const cellRef = `${numberToColumn(col)}${row}`;
      const cell = worksheet.getCell(cellRef);

      cell.border = {
        top: { style: row === startCell.row ? borderStyle : innerBorderStyle },
        left: { style: col === startColumn ? borderStyle : innerBorderStyle },
        bottom: { style: row === endCell.row ? borderStyle : innerBorderStyle },
        right: { style: col === endColumn ? borderStyle : innerBorderStyle }
      };
    }
  }
};
export class SpreadsheetCell {
  #initialColumn = null;
  #columnGenerator = null;
  #worksheet = null;
  #formatting = null;
  #currentPosition = null;

  constructor({ column, row, worksheet }) {
    if (!/^[A-Z]+$/.test(column))
      throw new Error("Formato de columna inválido");
    if (row < 1)
      throw new Error("La fila debe ser un número mayor o igual a 1");
    if (!worksheet)
      throw new Error("Se requiere un objeto worksheet válido");

    this.#initialColumn = column;
    this.#columnGenerator = generateColumnSequence(column);
    this.#worksheet = worksheet;
    this.#formatting = {};
    this.#currentPosition = { column: this.#columnGenerator.next().value, row };
  }

  columnToNumber(column) {
    return columnToNumber(column);
  }

  numberToColumn(number) {
    return numberToColumn(number);
  }

  get worksheet() {
    return this.#worksheet;
  }

  // Permite asignar null para liberar la referencia al worksheet
  set worksheet(value) {
    if (value !== null && typeof value.getCell !== 'function') {
      throw new Error("Debe proporcionar un worksheet válido o null");
    }
    this.#worksheet = value;
  }

  // Obtiene la celda actual basada en la posición (por ejemplo, "C3")
  get cell() {
    if (!this.#worksheet) throw new Error("Worksheet no definido");
    return this.#worksheet.getCell(`${this.#currentPosition.column}${this.#currentPosition.row}`);
  }

  applyFormatting(format = {}) {
    const effectiveFormat = { ...format, ... this.#formatting };
    for (const key in effectiveFormat) {
      this.cell[key] = { ...effectiveFormat[key] };
    }
  }

  overwriteFormatting(format) {
    for (const key in format) {
      this.cell[key] = { ...format[key] };
    }
  }

  moveToNextColumn(times = 1) {
    for (let i = 0; i < times; i++) {
      this.#currentPosition.column = this.#columnGenerator.next().value;
    }
  }

  moveToPreviousColumn(times = 1) {
    const currentNumber = columnToNumber(this.#currentPosition.column);
    const newNumber = currentNumber - times;
    if (newNumber < 1)
      throw new Error("No hay columnas anteriores a 'A'");
    this.#currentPosition.column = numberToColumn(newNumber);
    // Actualiza la columna inicial y reinicia el generador
    this.#initialColumn = this.#currentPosition.column;
    this.#columnGenerator = generateColumnSequence(this.#initialColumn);
  }

  moveToNextRow(times = 1) {
    this.#currentPosition.row += times;
    // Reinicia la secuencia a partir de la columna inicial
    this.#columnGenerator = generateColumnSequence(this.#initialColumn);
    this.#currentPosition.column = this.#columnGenerator.next().value;
  }

  moveToPreviousRow(times = 1) {
    const newRow = this.#currentPosition.row - times;
    if (newRow < 1)
      throw new Error("La fila no puede ser menor que 1");
    this.#currentPosition.row = newRow;
  }

  get position() {
    const self = this;
    return new Proxy(this.#currentPosition, {
      get(target, key) {
        return target[key];
      },
      set(target, key, value) {
        if (key === 'column') {
          self.setInitialColumn(value);
          target.column = self.#currentPosition.column;
          return true;
        }
        if (key === 'row') {
          if (value < 1) throw new Error("La fila no puede ser menor que 1");
          target.row = value;
          return true;
        }
        return true;
      },
      ownKeys(target) {
        return Reflect.ownKeys(target);
      },
      getOwnPropertyDescriptor(target, key) {
        return Object.getOwnPropertyDescriptor(target, key) || {
          configurable: true,
          enumerable: true,
          value: target[key]
        };
      }
    });
  }

  set position({ column, row }) {
    if (column) {
      this.setInitialColumn(column);
    }
    if (row) {
      if (row < 1) throw new Error("La fila no puede ser menor que 1");
      this.#currentPosition.row = row;
    }
  }

  setInitialColumn(newColumn) {
    if (!/^[A-Z]+$/.test(newColumn))
      throw new Error("Formato de columna inválido");
    this.#initialColumn = newColumn;
    this.#columnGenerator = generateColumnSequence(newColumn);
    this.#currentPosition.column = this.#columnGenerator.next().value;
  }

  moveTo({ columns = 0, rows = 0 }) {
    if (columns < 0) {
      this.moveToPreviousColumn(Math.abs(columns));
    } else {
      this.moveToNextColumn(columns);
    }
    if (rows < 0) {
      this.moveToPreviousRow(Math.abs(rows));
    } else {
      this.moveToNextRow(rows);
    }
  }

  get cellAddress() {
    return `${this.#currentPosition.column}${this.#currentPosition.row}`;
  }

  set cellAddress(value) {
    const [column, row] = value.match(/[A-Z]+|\d+/g);
    this.position = { column, row: Number(row) };
  }

  incrementColumnValue({ value, column }) {
    const targetColumn = column ?? this.#currentPosition.column;
    return incrementColumn(targetColumn, value);
  }

  setAndApplyFormat(value, format) {
    this.applyFormatting(format);
    this.cell.value = value;
  }

  setAndMoveToNextColumn(value, format) {
    this.setAndApplyFormat(value, format);
    this.moveToNextColumn();
  }

  setAndMoveToNextRow(value, format) {
    this.setAndApplyFormat(value, format);
    this.moveToNextRow();
  }

  setAndMoveToPreviousColumn(value, format) {
    this.setAndApplyFormat(value, format);
    this.moveToPreviousColumn();
  }

  setAndMoveToPreviousRow(value, format) {
    this.setAndApplyFormat(value, format);
    this.moveToPreviousRow();
  }

  set formatting(value) {
    this.#formatting = value;
  }

  get formatting() {
    return this.#formatting;
  }

  unsetFormatting() {
    this.#formatting = {};
  }

  addFormatting(format) {
    this.#formatting = { ...this.#formatting, ...format };
  }

  mergeCells({ columns = 0, rows = 0 }) {
    const targetColumn = this.incrementColumnValue({ value: columns });
    const targetRow = this.#currentPosition.row + rows;
    this.#worksheet.mergeCells(
      `${this.#currentPosition.column}${this.#currentPosition.row}:${targetColumn}${targetRow}`
    );
    this.#columnGenerator = generateColumnSequence(targetColumn);
    this.#currentPosition.column = this.#columnGenerator.next().value;
    this.#currentPosition.row = targetRow;
    console.log(`${this.#currentPosition.column}${this.#currentPosition.row}:${targetColumn}${targetRow}`);
  }

  fillRow({ values, format, exclude = [] }) {
    if (typeof values === "object" && !Array.isArray(values)) {
      values = Object.entries(values)
        .filter(([key]) => !exclude.includes(key))
        .map(([, value]) => value);
    }
    values.forEach((value) => {
      this.setAndMoveToNextColumn(value, format);
    });
  }

  fillColumns({ values, format }) {
    values.forEach((value) => {
      this.setAndMoveToNextRow(value, format);
    });
  }

  fillTable({ data, format, exclude = [] }) {
    data.forEach((row) => {
      this.fillRow({ values: row, format, exclude });
      this.moveToNextRow();
    });
  }

  set value(value) {
    this.cell.value = value;
  }

  get value() {
    return this.cell.value;
  }

  // Método auxiliar para liberar todas las referencias internas
  destroy() {
    this.#worksheet = null;
    this.#columnGenerator = null;
    this.#formatting = null;
    this.#currentPosition = null;
    this.#initialColumn = null;
  }

  clone() {
    return new SpreadsheetCell({
      column: this.#currentPosition.column,
      row: this.#currentPosition.row,
      worksheet: this.#worksheet,
    });
  }
}
