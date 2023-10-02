import Cell from "./Cell"
import SheetMemory from "./SheetMemory"
import { ErrorMessages } from "./GlobalDefinitions";

export class FormulaEvaluator {
  private _currentTokenIndex: number = 0;
  private _errorOccurred: boolean = false;
  private _errorMessage: string = "";
  private _currentFormula: FormulaType = [];
  private _lastResult: number = 0;
  private _sheetMemory: SheetMemory;
  private _result: number = 0;

  /**
   * Creates a new instance of the FormulaEvaluator class.
   * @param memory The SheetMemory object to use for evaluating formulas.
   */
  constructor(memory: SheetMemory) {
    this._sheetMemory = memory;
  }

  /**
   * Evaluates a given formula and returns the result.
   * The value of the formula is the value of the expression in the formula, 
   * using a recursive descent parser to evaluate the formula.
   * If formula has empty string, it should return an error; If formula has an error, 
   * it should return an error; Otherwise return the value of the formula.
   * @param formula - The formula to be evaluated.
   * @returns The result of the evaluated formula.
   */
  public evaluate(formula: FormulaType): number {
    this._currentFormula = formula;
    this._currentTokenIndex = 0;
    this._errorOccurred = false;
    this._errorMessage = "";

    if (formula.length === 0) {
      this._errorMessage = ErrorMessages.emptyFormula;
      return 0;
    }

    this._result = this.evaluateExpression();
    this._lastResult = this._result;
    if (this._currentTokenIndex !== formula.length && !this._errorOccurred) {
      this._errorMessage = ErrorMessages.invalidFormula;
      return 0;
    }

    return this._result;
  }

  /**
   * Calculates the value of an expression is the value of the 
   * first term plus or minus the value of the second term.
   * expression = term { ("+" | "-") term }.
   * @returns the value of the expression in the tokenized formula
   */
  private evaluateExpression(): number {
    let value = this.evaluateTerm();
    while (this.currentToken === '+' || this.currentToken === '-') {
      const operator = this.currentToken;
      this.nextToken();
      if (operator === '+') {
        value += this.evaluateTerm();
      } else if (operator === '-') {
        value -= this.evaluateTerm();
      }
    }
    return value;
  }

  /**
   * Evaluates a term by multiplying or dividing factors.
   * The value of a term is the value of the first factor times 
   * or divided by the value of the second factor.
   * term = factor { ("*" | "/") factor }.
   * @returns The result of the evaluated term.
   */
  private evaluateTerm(): number {
    let value = this.evaluateFactor();
    while (this.currentToken === '*' || this.currentToken === '/') {
      const operator = this.currentToken;
      this.nextToken();
      if (operator === '*') {
        value *= this.evaluateFactor();
      } else if (operator === '/') {
        const value2 = this.evaluateFactor();
        if (value2 === 0) {
          this._errorMessage = ErrorMessages.divideByZero;
          this._errorOccurred = true;
          this._lastResult = Infinity;
          return this._lastResult;
        }
        value /= value2;
      }
    }
    return value;
  }

  /**
   * Evaluates a factor in a formula expression.
   * The value of a factor is the value of the number or the value 
   * of the expression in the parentheses or the value of the cellReference.
   * factor = number | "(" expression ")" | cellReference returned by isCellReference.
   * @returns The numerical value of the factor.
   */
  private evaluateFactor(): number {
    let value = 0;
    let token;
    if (this.isNumber(this.currentToken)) {
      value = Number(this.currentToken);
      this.nextToken();
    } else if (this.isCellReference(this.currentToken)) {
      let [val, error] = this.getCellValue(this.currentToken);
      value = val;
      if (error) {
        this._errorMessage = error;
        this._errorOccurred = true;
      }
      this.nextToken();
    } else if (this.currentToken === '(') {
      token = this.nextToken();
      value = this.evaluateExpression();
      if (this.isCurrentToken(')')) {
        this.nextToken();
      } else {
        this._errorOccurred = true;
        this._errorMessage = ErrorMessages.invalidFormula;
      }
    } else {
      this._errorMessage = ErrorMessages.invalidFormula;
      this._errorOccurred = true;
      return 0;
    }
    return value;
  }

  /**
   * Advances to the next token from the formula being evaluated.
   */
  private nextToken(): void {
    this._currentTokenIndex++;
  }

  /**
   * @returns the current token
   */
  private get currentToken(): TokenType {
    return this._currentFormula[this._currentTokenIndex];
  }

  /**
   * Checks if the current token is of the specified type.
   * @param token The token type to check against the current token.
   * @returns True if the current token is of the specified type, false otherwise.
   */
  private isCurrentToken(token: TokenType): boolean {
    if (this.currentToken === token) {
      return true;
    } else {
      return false;
    }
  }

  /**
   * Gets the error message associated with the formula evaluation.
   * Returns an empty string if there are no errors.
   * @returns {string} The error message.
   */
  public get error(): string {
    return this._errorMessage
  }

  /**
   * Gets the result of the formula evaluation.
   * @returns The result of the formula evaluation as a number.
   */
  public get result(): number {
    return this._result;
  }

  /**
   * 
   * @param token 
   * @returns true if the toke can be parsed to a number
   */
  isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

  /**
   * Determines whether a given token is a valid cell reference.
   * @param token The token to check.
   * @returns True if the token is a valid cell reference, false otherwise.
   */
  isCellReference(token: TokenType): boolean {

    return Cell.isValidCellLabel(token);
  }

  /**
   * Retrieves the value and/or error of a referenced cell based on its token.
   * @param token The token representing the cell.
   * @returns A tuple containing the cell's value and error message (if any).
   */
  getCellValue(token: TokenType): [number, string] {

    let cell = this._sheetMemory.getCellByLabel(token);

    let formula = cell.getFormula();
    let error = cell.getError();

    // if the cell has an error return 0
    if (error !== "" && error !== ErrorMessages.emptyFormula) {
      return [0, error];
    }

    // if the cell formula is empty return 0
    if (formula.length === 0) {
      return [0, ErrorMessages.invalidCell];
    }

    let value = cell.getValue();
    return [value, ""];
  }
}

export default FormulaEvaluator;