import {Parser, ERROR_REF, error as isFormulaError} from 'hot-formula-parser';
import {arrayEach, arrayMap} from 'handsontable/helpers/array';
import localHooks from 'handsontable/mixins/localHooks';
import {getTranslator} from 'handsontable/utils/recordTranslator';
import {mixin} from 'handsontable/helpers/object';
import CellValue from './cell/value';
import CellReference from './cell/reference';
import {isFormulaExpression, toUpperCaseFormula} from './utils';
import Matrix from './matrix';
import AlterManager from './alterManager';

const STATE_UP_TO_DATE = 1;
const STATE_NEED_REBUILD = 2;
const STATE_NEED_FULL_REBUILD = 3;

/**
 * Sheet component responsible for whole spreadsheet calculations.
 *
 * @class Sheet
 * @util
 */
class Sheet {
  constructor(hot, dataProvider) {
    /**
     * Handsontable instance.
     *
     * @type {Core}
     */
    this.hot = hot;
    /**
     * Record translator for translating visual records into psychical and vice versa.
     *
     * @type {RecordTranslator}
     */
    this.t = getTranslator(this.hot);
    /**
     * Data provider for sheet calculations.
     *
     * @type {DataProvider}
     */
    this.dataProvider = dataProvider;
    /**
     * Instance of {@link https://github.com/handsontable/formula-parser}.
     *
     * @type {Parser}
     */
    this.parser = new Parser();
    /**
     * Instance of {@link Matrix}.
     *
     * @type {Matrix}
     */
    this.matrix = new Matrix(this.t);
    /**
     * Instance of {@link AlterManager}.
     *
     * @type {AlterManager}
     */
    this.alterManager = new AlterManager(this);
    /**
     * Cell object which indicates which cell is currently processing.
     *
     * @private
     * @type {null}
     */
    this._processingCell = null;
    /**
     * State of the sheet.
     *
     * @type {Number}
     * @private
     */
    this._state = STATE_NEED_FULL_REBUILD;

    this.parser.on('callCellValue', (...args) => this._onCallCellValue(...args));
    this.parser.on('callRangeValue', (...args) => this._onCallRangeValue(...args));
    this.alterManager.addLocalHook('afterAlter', (...args) => this._onAfterAlter(...args));
  }

  /**
   * Recalculate sheet.
   */
  recalculate() {
    switch (this._state) {
      case STATE_NEED_FULL_REBUILD:
        this.recalculateFull();
        break;
      case STATE_NEED_REBUILD:
        this.recalculateOptimized();
        break;
      default:
        break;
    }
  }

  /**
   * Recalculate sheet using optimized methods (fast recalculation).
   */
  recalculateOptimized(depth = 5) {
    const cells = this.matrix.getOutOfDateCells();
    let hasUncomputedFormulas = false;

    arrayEach(cells, (cellValue) => {
      if (this.matrix.getCellPrecedentsUpToDate(cellValue)) {
        const value = this.dataProvider.getSourceDataAtCell(cellValue.row, cellValue.column);

        if (isFormulaExpression(value)) {
          this.parseExpression(cellValue, value.substr(1));
        }
      } else {
        hasUncomputedFormulas = true;
      }
    });

    if (hasUncomputedFormulas && depth > 0) {
      this.recalculateOptimized(depth - 1);
    } else {
      this._state = STATE_UP_TO_DATE;
      this.runLocalHooks('afterRecalculate', cells, 'optimized');
    }
  }

  /**
   * Recalculate whole table by building dependencies from scratch (slow recalculation).
   */
  recalculateFull() {
    const cells = this.dataProvider.getSourceDataByRange();

    arrayEach(cells, (rowData, row) => {
      arrayEach(rowData, (value, column) => {
        if (isFormulaExpression(value)) {
          const cellValue = this.matrix.getCellAt(row, column) || new CellValue(row, column);
          this.parseExpression(cellValue, value.substr(1));
        }
      });
    });

    this._state = STATE_UP_TO_DATE;
    this.runLocalHooks('afterRecalculate', cells, 'full');
  }

  /**
   * Set predefined variable name which can be visible while parsing formula expression.
   *
   * @param {String} name Variable name.
   * @param {*} value Variable value.
   */
  setVariable(name, value) {
    this.parser.setVariable(name, value);
  }

  /**
   * Get variable name.
   *
   * @param {String} name Variable name.
   * @returns {*}
   */
  getVariable(name) {
    return this.parser.getVariable(name);
  }

  /**
   * Apply changes to the sheet.
   *
   * @param {Number} row Physical row index.
   * @param {Number} column Physical column index.
   * @param {*} newValue Current cell value.
   */
  applyChanges(row, column, newValue) {
    // Remove formula description for old expression
    // TODO: Move this to recalculate()
    const oldCellValue = this.matrix.getCellAt(row, column);
    const dependents = oldCellValue ? oldCellValue.getDependents() : [];
    this.matrix.remove({row, column});
    // ...and create new for new changed formula expression
    const cellValue = new CellValue(row, column);

    // copy over dependent values from old cell to new cell
    arrayEach(dependents, (dep) => cellValue.addDependent(dep));

    // TODO: Move this to recalculate()
    if (isFormulaExpression(newValue)) {
      this.parseExpression(cellValue, newValue.substr(1));
    } else {
      this.matrix.add(cellValue);
    }

    const deps = this.getCellDependencies(...this.t.toVisual(row, column));

    arrayEach(deps, (dep) => {
      dep.setState(CellValue.STATE_OUT_OFF_DATE);
    });

    this._state = STATE_NEED_REBUILD;
  }

  /**
   * Parse and evaluate formula for provided cell.
   *
   * @param {CellValue|Object} cellValue Cell value object.
   * @param {String} formula Value to evaluate.
   */
  parseExpression(cellValue, formula) {
    cellValue.setState(CellValue.STATE_COMPUTING);
    this._processingCell = cellValue;

    const {error, result} = this.parser.parse(toUpperCaseFormula(formula));

    if (isFormulaExpression(result)) {
      this.parseExpression(cellValue, result.substr(1));
    } else {
      cellValue.setValue(result);
      cellValue.setError(error);
      cellValue.setState(CellValue.STATE_UP_TO_DATE);
    }

    this.matrix.add(cellValue);
    this._processingCell = null;
  }

  /**
   * Get cell value object at specified physical coordinates.
   *
   * @param {Number} row Physical row index.
   * @param {Number} column Physical column index.
   * @returns {CellValue|undefined}
   */
  getCellAt(row, column) {
    return this.matrix.getCellAt(row, column);
  }

  /**
   * Get cell dependencies at specified physical coordinates.
   *
   * @param {Number} row Physical row index.
   * @param {Number} column Physical column index.
   * @returns {Array}
   */
  getCellDependencies(row, column) {
    return this.matrix.getDependencies({row, column});
  }

  /**
   * Listener for parser cell value.
   *
   * @private
   * @param {Object} cellCoords Cell coordinates.
   * @param {Function} done Function to call with valid cell value.
   */
  _onCallCellValue({row, column}, done) {
    if (!this.dataProvider.isInDataRange(row.index, column.index)) {
      throw Error(ERROR_REF);
    }

    const precedentCellRef = new CellReference(row, column);

    const dependentCellRef = new CellReference(this._processingCell.row, this._processingCell.column);

    this.matrix.registerCellRef(precedentCellRef);
    this.matrix.registerCellRef(dependentCellRef);
    this._processingCell.addPrecedent(precedentCellRef);

    let precedentCellValue;
    if (this.matrix.getCellAt(row.index, column.index)) {
      precedentCellValue = this.matrix.getCellAt(row.index, column.index);
      precedentCellValue.addDependent(dependentCellRef);
    } else {
      precedentCellValue = new CellValue(row, col);
      precedentCellValue.addDependent(dependentCellRef);
      this.matrix.add(precedentCellValue);
    }

    const cellValue = this.dataProvider.getRawDataAtCell(row.index, column.index);

    if (isFormulaError(cellValue) && precedentCellValue.hasError()) {
      throw Error(cellValue);
    }

    if (isFormulaExpression(cellValue)) {
      const {error, result} = this.parser.parse(cellValue.substr(1));

      if (error) {
        throw Error(error);
      }

      done(result);
    } else {
      done(cellValue);
    }
  }

  /**
   * Listener for parser cells (range) value.
   *
   * @private
   * @param {Object} startCell Cell coordinates (top-left corner coordinate).
   * @param {Object} endCell Cell coordinates (bottom-right corner coordinate).
   * @param {Function} done Function to call with valid cells values.
   */
  _onCallRangeValue({row: startRow, column: startColumn}, {row: endRow, column: endColumn}, done) {
    const cellValues = this.dataProvider.getRawDataByRange(startRow.index, startColumn.index, endRow.index, endColumn.index);

    const mapRowData = (rowData, rowIndex) => arrayMap(rowData, (cellData, columnIndex) => {
      const rowCellCoord = startRow.index + rowIndex;
      const columnCellCoord = startColumn.index + columnIndex;

      if (!this.dataProvider.isInDataRange(rowCellCoord, columnCellCoord)) {
        throw Error(ERROR_REF);
      }

      const precedentCellRef = new CellReference(rowCellCoord, columnCellCoord);

      const dependentCellRef = new CellReference(this._processingCell.row, this._processingCell.column);

      this.matrix.registerCellRef(precedentCellRef);
      this.matrix.registerCellRef(dependentCellRef);
      this._processingCell.addPrecedent(precedentCellRef);

      let precedentCellValue;
      if (this.matrix.getCellAt(rowCellCoord, columnCellCoord)) {
        precedentCellValue = this.matrix.getCellAt(rowCellCoord, columnCellCoord);
        precedentCellValue.addDependent(dependentCellRef);
      } else {
        precedentCellValue = new CellValue(rowCellCoord, columnCellCoord);
        precedentCellValue.addDependent(dependentCellRef);
        this.matrix.add(precedentCellValue);
      }

      if (isFormulaError(cellData)) {
        const computedCell = this.matrix.getCellAt(cell.row, cell.column);

        if (computedCell && computedCell.hasError()) {
          throw Error(cellData);
        }
      }

      if (isFormulaExpression(cellData)) {
        const {error, result} = this.parser.parse(cellData.substr(1));

        if (error) {
          throw Error(error);
        }

        cellData = result;
      }

      return cellData;
    });

    const calculatedCellValues = arrayMap(cellValues, (rowData, rowIndex) => mapRowData(rowData, rowIndex));

    done(calculatedCellValues);
  }

  /**
   * On after alter sheet listener.
   *
   * @private
   */
  _onAfterAlter() {
    this.recalculateOptimized();
  }

  /**
   * Destroy class.
   */
  destroy() {
    this.hot = null;
    this.t = null;
    this.dataProvider.destroy();
    this.dataProvider = null;
    this.alterManager.destroy();
    this.alterManager = null;
    this.parser = null;
    this.matrix.reset();
    this.matrix = null;
  }
}

mixin(Sheet, localHooks);

export default Sheet;
