export class GoogleSpreadsheetFormulaError {
  public type: string;
  public message: string;
  
  constructor(errorInfo: { type: string, message: string }) {
    this.type = errorInfo.type;
    this.message = errorInfo.message;
  }
}