const crypto = require('crypto');

class MockLock {
  constructor() {
    this.hasLock = false;
  }
  tryLock(timeInMillis) {
    this.hasLock = true;
    return true;
  }
  waitLock(timeInMillis) {
    this.hasLock = true;
  }
  releaseLock() {
    this.hasLock = false;
  }
  hasLock() {
    return this.hasLock;
  }
}

class MockProperties {
  constructor() {
    this.props = {};
  }
  getProperty(key) {
    return this.props[key] || null;
  }
  setProperty(key, value) {
    this.props[key] = String(value);
    return this;
  }
  getProperties() {
    return { ...this.props };
  }
  setProperties(properties) {
    for (const key in properties) {
      this.props[key] = String(properties[key]);
    }
    return this;
  }
  deleteProperty(key) {
    delete this.props[key];
    return this;
  }
  deleteAllProperties() {
    this.props = {};
    return this;
  }
}

class MockSpreadsheetApp {
  constructor() {
    this.activeSpreadsheet = new MockSpreadsheet();
  }
  getActiveSpreadsheet() {
    return this.activeSpreadsheet;
  }
  newDataValidation() {
    return new MockDataValidationBuilder();
  }
  reset() {
    this.activeSpreadsheet = new MockSpreadsheet();
  }
  getUi() {
    return {
        createMenu: (name) => ({
             addItem: (label, func) => ({
                 addToUi: () => {}
             })
        }),
        alert: () => {},
        ButtonSet: { OK: 1 }
    };
  }
}

class MockSpreadsheet {
  constructor() {
    this.sheets = [];
  }
  getSheetByName(name) {
    return this.sheets.find(s => s.name === name) || null;
  }
  getSheets() {
    return this.sheets;
  }
  insertSheet(name) {
    if (this.getSheetByName(name)) {
      throw new Error(`Sheet ${name} already exists.`);
    }
    const sheet = new MockSheet(name);
    this.sheets.push(sheet);
    return sheet;
  }
}

class MockSheet {
  constructor(name) {
    this.name = name;
    this.data = []; // 2D array [row][col], 0-indexed
    this.validations = []; // parallel array
    this.formulas = [];
    this.notes = [];
    this.backgrounds = [];
  }
  getName() {
    return this.name;
  }
  getDataRange() {
    return this.getRange(1, 1, Math.max(1, this.data.length), Math.max(1, this.data[0]?.length || 1));
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    return new MockRange(this, row, col, numRows, numCols);
  }
  getMaxRows() {
    return this.data.length;
  }
  getLastRow() {
    return this.data.length;
  }
  getLastColumn() {
    return this.data.length > 0 ? this.data[0].length : 0;
  }
  insertColumnAfter(afterPosition) {
    for (let r = 0; r < this.data.length; r++) {
      this.data[r].splice(afterPosition, 0, "");
      if (this.validations[r]) this.validations[r].splice(afterPosition, 0, null);
      if (this.formulas[r]) this.formulas[r].splice(afterPosition, 0, "");
      if (this.notes[r]) this.notes[r].splice(afterPosition, 0, "");
      if (this.backgrounds[r]) this.backgrounds[r].splice(afterPosition, 0, "");
    }
    return this;
  }
  insertColumnsAfter(afterPosition, howMany) {
    for (let i = 0; i < howMany; i++) {
        this.insertColumnAfter(afterPosition + i);
    }
    return this;
  }
  appendRow(rowContents) {
      this.data.push([...rowContents]);
      this.validations.push(new Array(rowContents.length).fill(null));
      this.formulas.push(new Array(rowContents.length).fill(""));
      this.notes.push(new Array(rowContents.length).fill(""));
      this.backgrounds.push(new Array(rowContents.length).fill(""));
      return this;
  }
  clear() {
      this.data = [];
      this.validations = [];
      this.formulas = [];
      this.notes = [];
      this.backgrounds = [];
  }
}

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row; // 1-indexed
    this.col = col; // 1-indexed
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValues() {
    const vals = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowVals = [];
      for (let c = 0; c < this.numCols; c++) {
        rowVals.push(this.sheet.data[this.row - 1 + r]?.[this.col - 1 + c] ?? "");
      }
      vals.push(rowVals);
    }
    return vals;
  }
  setValues(values) {
    if (values.length !== this.numRows || (values[0] && values[0].length !== this.numCols)) {
       throw new Error("setValues size mismatch");
    }
    for (let r = 0; r < this.numRows; r++) {
      if (!this.sheet.data[this.row - 1 + r]) {
         this.sheet.data[this.row - 1 + r] = [];
      }
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.data[this.row - 1 + r][this.col - 1 + c] = values[r][c];
      }
    }
    return this;
  }
  setValue(value) {
    return this.setValues([[value]]);
  }
  getValue() {
      return this.getValues()[0][0];
  }
  setDataValidations(validations) {
    for (let r = 0; r < this.numRows; r++) {
      if (!this.sheet.validations[this.row - 1 + r]) {
         this.sheet.validations[this.row - 1 + r] = [];
      }
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.validations[this.row - 1 + r][this.col - 1 + c] = validations[r][c];
      }
    }
    return this;
  }
  setDataValidation(validation) {
    const v = [];
    for (let r = 0; r < this.numRows; r++) {
        v.push(new Array(this.numCols).fill(validation));
    }
    return this.setDataValidations(v);
  }
  getDataValidation() {
      const r = this.row - 1;
      const c = this.col - 1;
      if (this.sheet.validations[r] && this.sheet.validations[r][c]) {
          return this.sheet.validations[r][c];
      }
      return null;
    }
    getDataValidations() {
    const vals = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowVals = [];
      for (let c = 0; c < this.numCols; c++) {
        rowVals.push(this.sheet.validations[this.row - 1 + r]?.[this.col - 1 + c] ?? null);
      }
      vals.push(rowVals);
    }
    return vals;
  }
  setFormulas(formulas) {
    for (let r = 0; r < this.numRows; r++) {
      if (!this.sheet.formulas[this.row - 1 + r]) {
         this.sheet.formulas[this.row - 1 + r] = [];
      }
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.formulas[this.row - 1 + r][this.col - 1 + c] = formulas[r][c];
      }
    }
    return this;
  }
  getFormulas() {
    const vals = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowVals = [];
      for (let c = 0; c < this.numCols; c++) {
        rowVals.push(this.sheet.formulas[this.row - 1 + r]?.[this.col - 1 + c] ?? "");
      }
      vals.push(rowVals);
    }
    return vals;
  }
  setNotes(notes) {
      for (let r = 0; r < this.numRows; r++) {
      if (!this.sheet.notes[this.row - 1 + r]) {
         this.sheet.notes[this.row - 1 + r] = [];
      }
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.notes[this.row - 1 + r][this.col - 1 + c] = notes[r][c];
      }
    }
    return this;
  }
  getNotes() {
    const vals = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowVals = [];
      for (let c = 0; c < this.numCols; c++) {
        rowVals.push(this.sheet.notes[this.row - 1 + r]?.[this.col - 1 + c] ?? "");
      }
      vals.push(rowVals);
    }
    return vals;
  }
  getNumberFormat() {
      if (this.sheet.formats && this.sheet.formats[this.row-1] && this.sheet.formats[this.row-1][this.col-1]) {
          return this.sheet.formats[this.row-1][this.col-1];
      }
      return '';
    }
    setNumberFormat(format) {
      if (!this.sheet.formats) this.sheet.formats = [];
      for (let i = 0; i < this.numRows; i++) {
        for (let j = 0; j < this.numCols; j++) {
          const r = this.row - 1 + i;
          const c = this.col - 1 + j;
          if (!this.sheet.formats[r]) this.sheet.formats[r] = [];
          this.sheet.formats[r][c] = format;
        }
      }
      return this;
    }
  setBackgrounds(backgrounds) {
    for (let r = 0; r < this.numRows; r++) {
      if (!this.sheet.backgrounds[this.row - 1 + r]) {
         this.sheet.backgrounds[this.row - 1 + r] = [];
      }
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.backgrounds[this.row - 1 + r][this.col - 1 + c] = backgrounds[r][c];
      }
    }
    return this;
  }
  getBackgrounds() {
    const vals = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowVals = [];
      for (let c = 0; c < this.numCols; c++) {
        rowVals.push(this.sheet.backgrounds[this.row - 1 + r]?.[this.col - 1 + c] ?? "");
      }
      vals.push(rowVals);
    }
    return vals;
  }
  offset(rowOffset, colOffset, numRows = this.numRows, numCols = this.numCols) {
     return new MockRange(this.sheet, this.row + rowOffset, this.col + colOffset, numRows, numCols);
  }
}

class MockDataValidation {
        constructor(builder) {
            this.criteriaType = builder.type || '';
            this.args = builder.args || [];
            this.helpText = builder.helpText || '';
            this._allowInvalid = builder._allowInvalid !== false;
        }
        getCriteriaType() { return this.criteriaType; }
        getCriteriaValues() { return this.args; }
        getHelpText() { return this.helpText; }
        getAllowInvalid() { return this._allowInvalid; }
    }

    class MockDataValidationBuilder {
  constructor() {
    this.type = null;
    this.args = [];
  }
  requireFormulaSatisfied(formula) { this.type = 'FORMULA'; this.args = [formula]; return this; }
      setHelpText(text) { this.helpText = text; return this; }
      getCriteriaType() { return this.type || ''; }
      getCriteriaValues() { return this.args || []; }
      getHelpText() { return this.helpText || ''; }
      getAllowInvalid() { return this._allowInvalid !== false; }
      requireCheckbox() {
    this.type = 'CHECKBOX';
    return this;
  }
  requireValueInList(values, showDropdown) {
    this.type = 'VALUE_IN_LIST';
    this.args = [values, showDropdown];
    return this;
  }
  requireNumberGreaterThanOrEqualTo(value) {
    this.type = 'NUMBER_GREATER_THAN_OR_EQUAL_TO';
    this.args = [value];
    return this;
  }
  requireNumberBetween(start, end) {
      this.type = 'NUMBER_BETWEEN';
      this.args = [start, end];
      return this;
  }
  allowInvalid(allow) {
    this.allowInvalid = allow;
    return this;
  }
  build() {
    return { type: this.type, args: this.args, allowInvalid: this.allowInvalid };
  }
}

class MockUtilities {
  constructor() {
    this.DigestAlgorithm = {
      SHA_256: 'SHA_256'
    };
    this.Charset = {
      UTF_8: 'UTF_8'
    };
  }
  computeDigest(algorithm, value, charset) {
    // Return mock byte array
    const hash = crypto.createHash('sha256').update(value).digest();
    return Array.from(hash);
  }
  base64Encode(bytes) {
      return Buffer.from(bytes).toString('base64');
  }
  base64Decode(str) {
      return Array.from(Buffer.from(str, 'base64'));
  }
  getUuid() {
      return crypto.randomUUID();
  }
}

class MockHtmlService {
  createHtmlOutput(content) {
    return {
        content: content,
        setTitle(title) { this.title = title; return this; },
        setXFrameOptionsMode(mode) { this.xFrameOptionsMode = mode; return this; },
        addMetaTag(name, content) { return this; }
    };
  }
  createTemplateFromFile(filename) {
    return {
        filename,
        evaluate: () => this.createHtmlOutput(`<!-- mock evaluated ${filename} -->`)
    };
  }
}
MockHtmlService.XFrameOptionsMode = { ALLOWALL: 'ALLOWALL' };

class MockLogger {
  log(...args) {
    // console.log(...args); // Keep silent during normal tests
  }
}

global.SpreadsheetApp = new MockSpreadsheetApp();
global.PropertiesService = {
  getScriptProperties: () => global.mockScriptProperties
};
global.LockService = {
  getScriptLock: () => global.mockScriptLock
};
global.Utilities = new MockUtilities();
global.HtmlService = new MockHtmlService();
global.Logger = new MockLogger();
global.mockScriptProperties = new MockProperties();
global.mockScriptLock = new MockLock();
global.console.log = (...args) => {}; // suppress unless debugging

module.exports = {
  resetMock: () => {
    global.SpreadsheetApp.reset();
    global.mockScriptProperties.deleteAllProperties();
    global.mockScriptLock.releaseLock();
  }
};
