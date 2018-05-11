import {arrayEach, arrayFilter, arrayReduce} from 'handsontable/helpers/array';
import CellValue from './cell/value';

/**
 * This component is responsible for storing all calculated cells which contain formula expressions (CellValue) and
 * register for all cell references (CellReference).
 *
 * CellValue is an object which represents a formula expression. It contains a calculated value of that formula,
 * an error if applied and cell references. Cell references are CellReference object instances which represent a cell
 * in a spreadsheet. One CellReference can be assigned to multiple CellValues as a precedent cell. Each cell
 * modification triggers a search through CellValues that are dependent of the CellReference. After
 * the match, the cells are marked as 'out of date'. In the next render cycle, all CellValues marked with
 * that state are recalculated.
 *
 * @class Matrix
 * @util
 */
class Matrix {
  constructor(recordTranslator) {
    /**
     * Record translator for translating visual records into psychical and vice versa.
     *
     * @type {RecordTranslator}
     */
    this.t = recordTranslator;
    /**
     * List of all cell values with theirs precedents.
     *
     * @type {Map}
     */
    this.data = new Map();
    /**
     * List of all created and registered cell references.
     *
     * @type {Array}
     */
    this.cellReferences = [];

    /**
     * List of all created and registered dependent containers.
     *
     * @type {Map}
     */
    this.dependentContainers = new Map();
  }

  /**
   * Get cell value at given row and column index.
   *
   * @param {Number} row Physical row index.
   * @param {Number} column Physical column index.
   * @returns {CellValue|null} Returns CellValue instance or `null` if cell not found.
   */
  getCellAt(row, column) {
    return this.data.get(`${row}, ${column}`) || null;
  }

  /**
   * Get dependent container at given row and column index.
   *
   * @param {Number} row Physical row index.
   * @param {Number} column Physical column index.
   * @returns {DependentContainer|null} Returns DependentContainer instance or `null` if cell not found.
   */
  getDependentContainerAt(row, column) {
    return this.dependentContainers.get(`${row}, ${column}`) || null;
  }

  /**
   * Get all out of date cells.
   *
   * @returns {Array}
   */
  getOutOfDateCells() {
    return arrayFilter(this.data.values(), (cell) => cell.isState(CellValue.STATE_OUT_OFF_DATE));
  }

  /**
   * Get all out of date cells.
   *
   * @returns {Array}
   */
  setCellsOutOfDate() {
    arrayEach(this.data.values(), (cell) => cell.setState(CellValue.STATE_OUT_OFF_DATE));
  }

  getCellPrecedentsUpToDate(cellValue) {
    return arrayFilter(cellValue.getPrecedents(), ({row, column}) => {
      const cell = this.getCellAt(row, column);
      return cell && !cell.isState(CellValue.STATE_UP_TO_DATE);
    }
    ).length === 0;
  }

  /**
   * Add cell value to the collection.
   *
   * @param {CellValue|Object} cellValue Cell value object.
   */
  add(cellValue) {
    const { row, column } = cellValue;
    if (!this.getCellAt(row, column)) {
      this.data.set(`${row}, ${column}`, cellValue);
    }
  }

  /**
   * Remove cell value from the collection.
   *
   * @param {CellValue|Object|Array} cellValue Cell value object.
   */
  remove(cellValue) {
    const isArray = Array.isArray(cellValue);
    if (isArray) {
      arrayEach(cellValue, ({row, column}) => {
        this.data.delete(`${row}, ${column}`);
      });
    } else {
      this.data.delete(`${cellValue.row}, ${cellValue.column}`);
    }
  }

  /**
   * Creates new maps for cellValues and dependentContainers that correspond to new locations
   *
   * @param {number} locatio where the change starts
   */
  translateCells(start, translate) {
    // create new map for cellValues at new location
    const axis = translate[0] > 0 ? 'row' : 'col';
    this.data = arrayReduce(this.data.values(), (map, cell) => {
      if ((axis === 'row' && cell.row >= start) || (axis === 'col' && cell.column >= start)) {
        cell.translateTo(...translate);
        cell.setState(CellValue.STATE_OUT_OFF_DATE);
      }

      map.set(`${cell.row}, ${cell.column}`, cell);
      return map;
    }, new Map());

    // create new map for dependentContainers at new location
    this.dependentContainers = arrayReduce(this.dependentContainers.values(), (map, cell) => {
      if ((axis === 'row' && cell.row >= start) || (axis === 'col' && cell.column >= start)) {
        cell.translateTo(...translate);
      }

      map.set(`${cell.row}, ${cell.column}`, cell);
      return map;
    }, new Map());
  }

  /**
   * Get cell dependencies using visual coordinates.
   *
   * @param {Object} cellCoord Visual cell coordinates object.
   */
  getDependencies({ row, column }) {
    const container = this.getDependentContainerAt(row, column);
    return container ? container.getDependents().map((dep) => {
      const visualCoords = this.t.toVisual(dep.row, dep.column);
      return this.getCellAt(...visualCoords);
    }).filter((dep) => !!dep) : [];
  }

  /**
   * Register cell reference to the collection.
   *
   * @param {CellReference|Object} cellReference Cell reference object.
   */
  registerCellRef(cellReference) {
    if (!arrayFilter(this.cellReferences, (cell) => cell.isEqual(cellReference)).length) {
      this.cellReferences.push(cellReference);
    }
  }

  /**
   * Register dependent container
   *
   * @param {DependentContainer|Object} cellReference Cell reference object.
   */
  registerDependentContainer(dependentContainer) {
    const { row, column } = dependentContainer;
    if (!this.getDependentContainerAt(row, column)) {
      this.dependentContainers.set(`${row}, ${column}`, dependentContainer);
    }
  }

  /**
   * Remove cell references from the collection.
   *
   * @param {Object} start Start visual coordinate.
   * @param {Object} end End visual coordinate.
   * @returns {Array} Returns removed cell references.
   */
  removeCellRefsAtRange({row: startRow, column: startColumn}, {row: endRow, column: endColumn}) {
    const removed = [];

    const rowMatch = (cell) => (startRow === void 0 ? true : cell.row >= startRow && cell.row <= endRow);
    const colMatch = (cell) => (startColumn === void 0 ? true : cell.column >= startColumn && cell.column <= endColumn);

    this.cellReferences = arrayFilter(this.cellReferences, (cell) => {
      if (rowMatch(cell) && colMatch(cell)) {
        removed.push(cell);
        this.dependentContainers.delete(`${cell.row}, ${cell.column}`);

        return false;
      }

      return true;
    });

    return removed;
  }

  /**
   * Reset matrix data.
   */
  reset() {
    this.data = new Map();
    this.cellReferences.length = 0;
  }
}

export default Matrix;
