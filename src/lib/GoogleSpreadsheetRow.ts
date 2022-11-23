import GoogleSpreadsheetWorksheet from './GoogleSpreadsheetWorksheet';
import { columnToLetter } from './utils';


// TODO: add type for possible row values (currently any)

export default class GoogleSpreadsheetRow<T extends Record<string, any> = Record<string, any>> {  
  constructor(
    /** parent GoogleSpreadsheetWorksheet instance */
    readonly _worksheet: GoogleSpreadsheetWorksheet,
    /** the A1 row (1-indexed) */
    readonly _rowNumber: number,
    /** raw underlying data for row */
    private _rawData: any[],
  ) {
    return this;
  }

  private _deleted = false;
  get deleted() { return this._deleted; }
  get rowNumber() { return this._rowNumber; }
  get a1Range() {
    return [
      this._worksheet.a1SheetName,
      '!',
      `A${this._rowNumber}`,
      ':',
      `${columnToLetter(this._worksheet.headerValues.length)}${this._rowNumber}`,
    ].join('');
  }

  /** get row's value of specific cell (by header key) */
  get(key: keyof T) {
    const index = this._worksheet.headerValues.indexOf(key as string);
    return this._rawData[index];
  }
  /** set row's value of specific cell (by header key) */
  set<K extends keyof T>(key: K, val: T[K]) { 
    const index = this._worksheet.headerValues.indexOf(key as string);
    this._rawData[index] = val;
  }
  /** set multiple values in the row at once from an object */
  assign(obj: T) {
    for (const key in obj) this.set(key, obj[key]);
  }

  /** return raw object of row data */
  toObject() {
    const o: Partial<T> = {};
    for (let i = 0; i < this._worksheet.headerValues.length; i++) {
      const key: keyof T = this._worksheet.headerValues[i];
      if (!key) continue;
      o[key] = this._rawData[i];
    }
    return o;
  }

  /** save row values */
  async save(options?: { raw?: boolean }) {
    if (this._deleted) throw new Error('This row has been deleted - call getRows again before making updates.');

    const response = await this._worksheet._spreadsheet.sheetsApi.request({
      method: 'put',
      url: `/values/${encodeURIComponent(this.a1Range)}`,
      params: {
        valueInputOption: options?.raw ? 'RAW' : 'USER_ENTERED',
        includeValuesInResponse: true,
      },
      data: {
        range: this.a1Range,
        majorDimension: 'ROWS',
        values: [this._rawData],
      },
    });
    this._rawData = response.data.updatedData.values[0];
  }

  /** delete this row */
  async delete() {
    if (this._deleted) throw new Error('This row has been deleted - call getRows again before making updates.');

    const result = await this._worksheet._makeSingleUpdateRequest('deleteRange', {
      range: {
        sheetId: this._worksheet.sheetId,
        startRowIndex: this._rowNumber - 1, // this format is zero indexed, because of course...
        endRowIndex: this._rowNumber,
      },
      shiftDimension: 'ROWS',
    });
    this._deleted = true;
    return result;
  }
}
