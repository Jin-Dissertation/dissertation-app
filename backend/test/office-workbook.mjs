// Minimal behavioral Excel model for importer tests. Tenant/API execution still
// requires an institutional smoke test; this does not emulate Excel itself.
export class Workbook {
  tables = new Map();
  sheets = new Map();
  failWrite = null;
  formulas = [];
  getTable(name) { return this.tables.get(name); }
  getWorksheet(name) { return this.sheets.get(name); }
  addWorksheet(name) {
    const sheet = {
      getRangeByIndexes: () => {
        const range = { values: [], setValues: values => { range.values = values; } };
        return range;
      },
      addTable: range => new Table(this, range.values[0])
    };
    this.sheets.set(name, sheet);
    return sheet;
  }
  write(name, action) {
    if (this.failWrite) this.failWrite(name, action);
  }
  cell(value) {
    if (typeof value !== "string") return value;
    if (value.length > 32767) throw new Error("Excel cell limit");
    if (value.startsWith("'")) return value.slice(1);
    if (/^[\s\u0000-\u001f]*[=+\-@]/.test(value)) this.formulas.push(value);
    return value;
  }
}

class Table {
  rows = [];
  constructor(workbook, headers) { this.workbook = workbook; this.headers = headers; }
  setName(name) { this.name = name; this.workbook.tables.set(name, this); }
  getName() { return this.name; }
  getRowCount() { return this.rows.length; }
  getHeaderRowRange() { return { getValues: () => [this.headers] }; }
  getRangeBetweenHeaderAndTotal() {
    return {
      getValues: () => structuredClone(this.rows),
      getRow: index => ({
        getValues: () => [structuredClone(this.rows[index])],
        setValues: values => {
          this.workbook.write(this.name, "setValues");
          this.rows[index] = values[0].map(v => this.workbook.cell(v));
        }
      })
    };
  }
  getColumnByName(name) {
    const column = this.headers.indexOf(name);
    if (column < 0) return undefined;
    return { getRangeBetweenHeaderAndTotal: () => ({ getValues: () => this.rows.map(row => [row[column]]) }) };
  }
  addRows(index, rows) {
    this.workbook.write(this.name, "addRows");
    this.rows.splice(index < 0 ? this.rows.length : index, 0, ...rows.map(row => row.map(v => this.workbook.cell(v))));
  }
  deleteRowsAt(index, count) {
    this.workbook.write(this.name, "deleteRowsAt");
    this.rows.splice(index, count);
  }
}
