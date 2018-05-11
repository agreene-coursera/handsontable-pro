import {ERROR_REF} from 'hot-formula-parser';
import {arrayFilter} from 'handsontable/helpers/array';
import BaseCell from './_base';

/**
 * Class which indicates formula expression precedents cells at specified cell
 * coordinates (CellValue). This object uses visual cell coordinates.
 *
 * @class CellReference
 * @util
 */
class DependentContainer extends BaseCell {
  constructor(row, column) {
    super(row, column);
    /**
     * List of dependent cells.
     *
     * @type {Array}
     */
    this.dependents = [];
  }

  /**
   * Add dependent cell to the collection.
   *
   * @param {CellReference} cellReference Cell reference object.
   */
  addDependent(cellReference) {
    if (this.isEqual(cellReference)) {
      throw Error(ERROR_REF);
    }
    if (!this.hasDependent(cellReference)) {
      this.dependents.push(cellReference);
    }
  }

  /**
   * Remove dependent cell from the collection.
   *
   * @param {CellReference} cellReference Cell reference object.
   */
  removeDependent(cellReference) {
    if (this.isEqual(cellReference)) {
      throw Error(ERROR_REF);
    }
    this.precedents = arrayFilter(this.dependents, (cell) => !cell.isEqual(cellReference));
  }

  /**
   * Get dependent cells.
   *
   * @returns {Array}
   */
  getDependents() {
    return this.dependents;
  }

  /**
   * Clear all dependent cells.
   */
  clearDependents() {
    this.dependents.length = 0;
  }

  /**
   * Check if cell value has dependent cells.
   *
   * @returns {Boolean}
   */
  hasDependents() {
    return this.dependents.length > 0;
  }

  /**
   * Check if cell reference is dependent on this cell.
   *
   * @param {CellReference} cellReference Cell reference object.
   * @returns {Boolean}
   */

  hasDependent(cellReference) {
    return arrayFilter(this.dependents, (cell) => cell.isEqual(cellReference)).length > 0;
  }
}

export default DependentContainer;
