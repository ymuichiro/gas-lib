/**
 * Spreadsheetを操作する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function SpreadsheetLib(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  function openSheet(sheetName: string) {
    const sheet = ss.getSheetByName(sheetName);
    if (TypeGuard.isNull(sheet)) {
      throw new Error(`sheet "${sheetName}" は Spreadsheet ${ss.getId()} に存在しません`);
    }
    return { sheet };
  }

  // SpreadsheetLib return
  return { openSheet };
}

/**
 * 入力された文字列をMD5形式でHash化する
 * @param {string} text
 * @returns {string} hash
 */
export function createMD5HashKey(text: string): string {
  const md5_byte = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text, Utilities.Charset.UTF_8);
  let hash_key = '';
  md5_byte.forEach(byte => (hash_key += byte < 0 ? (byte += 256).toString(16) : byte.toString(16)));
  return hash_key;
}

class TypeGuard {
  /**
   * 与えられた引数が undefined か否かを返す
   * @param v 判定対象の値
   * @returns {v is undefined} v is undefined
   */
  static isUndefined<T>(v: T | undefined): v is undefined {
    return v === undefined;
  }
  /**
   * 与えられた引数が null か否かを返す
   * @param v 判定対象の値
   * @returns {v is null} v is null
   */
  static isNull<T>(v: T | null): v is null {
    return v === null;
  }
  /**
   * 与えられた引数の長さが 0 か否かを返す
   * @param v 判定対象の配列
   * @returns {boolean} boolean
   */
  static isZeroLength<T>(v: T[]): boolean {
    return v.length === 0;
  }
  /**
   * 与えられた日付オブジェクトがInvalid Dateか否かを返す
   * @param v 判定対象のDate Object
   * @returns {boolean} boolean
   */
  static isInvalidDate(d: Date): boolean {
    return d.toString() === 'Invalid Date';
  }
  /**
   * 与えられた数値がNaNか否かを返す
   * @param v 判定対象の数値
   * @returns {boolean} boolean
   */
  static isNan(n: number): boolean {
    return n.toString() === 'NaN';
  }
}

interface GoogleSheetsOpenEvents {
  authMode: GoogleAppsScript.Script.AuthMode;
  source: GoogleAppsScript.Spreadsheet.Spreadsheet;
  triggerUid: string;
  user: string;
}

type SheetChangeType =
  | 'EDIT'
  | 'INSERT_ROW'
  | 'INSERT_COLUMN'
  | 'REMOVE_ROW'
  | 'REMOVE_COLUMN'
  | 'INSERT_GRID'
  | 'REMOVE_GRID'
  | 'FORMAT'
  | 'OTHER';

interface GoogleSheetsChangeEvents {
  authMode: GoogleAppsScript.Script.AuthMode;
  changeType: SheetChangeType;
  triggerUid: string;
  user: string;
}

interface GoogleSheetsEditEvents {
  authMode: GoogleAppsScript.Script.AuthMode;
  oldValue: any;
  range: GoogleAppsScript.Spreadsheet.Range;
  source: GoogleAppsScript.Spreadsheet.Spreadsheet;
  value: any;
  triggerUid: string;
  user: string;
}

interface GoogleSheetSubmitEvents {
  authMode: GoogleAppsScript.Script.AuthMode;
  /** An object containing the question names and values from the form submission. */
  namedValues: { [key: string]: any };
  range: GoogleAppsScript.Spreadsheet.Range;
  triggerUid: string;
  values: any[];
}

interface GoogleSlideOpenEvents {
  authMode: GoogleAppsScript.Script.AuthMode;
  source: GoogleAppsScript.Slides.Presentation;
  user: string;
}

interface GoogleFormOpenEvents {
  authModde: GoogleAppsScript.Script.AuthMode;
  source: GoogleAppsScript.Forms.Form;
  triggerUid: string;
  user: string;
}

interface GoogleFormSubmitEvents {
  authMode: GoogleAppsScript.Script.AuthMode;
  response: GoogleAppsScript.Forms.FormResponse;
  source: GoogleAppsScript.Forms.Form;
  triggerUid: string;
}

interface GoogleCalenderUpdateEvents {
  authMode: GoogleAppsScript.Script.AuthMode;
  calenderId: string;
  triggerUid: string;
}

interface GoogleCalenderTimeDrivenEvents {
  authMode: GoogleAppsScript.Script.AuthMode;
  'day-of-month': number;
  'day-of-week': number;
  hour: number;
  minute: number;
  month: number;
  second: number;
  timezone: string;
  triggerUid: string;
  'week-of-year': string;
}

class EventAddType {
  /** @returns {GoogleFormOpenEvents} */
  static sheetOpenEvents(v: any): GoogleSheetsOpenEvents {
    return v as GoogleSheetsOpenEvents;
  }
  /** @returns {GoogleSheetsChangeEvents} */
  static sheetChangeEvents(v: any): GoogleSheetsChangeEvents {
    return v as GoogleSheetsChangeEvents;
  }
  /** @returns {GoogleSheetsEditEvents} */
  static sheetEditEvents(v: any): GoogleSheetsEditEvents {
    return v as GoogleSheetsEditEvents;
  }
  /** @returns {GoogleSheetsSubmitEvents} */
  static sheetSubmitEvents(v: any): GoogleSheetSubmitEvents {
    return v as GoogleSheetSubmitEvents;
  }
  /** @returns {GoogleSlideOpenEvents} */
  static slideOpenEvents(v: any): GoogleSlideOpenEvents {
    return v as GoogleSlideOpenEvents;
  }
  /** @returns {GoogleFormOpenEvents} */
  static formOpenEvents(v: any): GoogleFormOpenEvents {
    return v as GoogleFormOpenEvents;
  }
  /** @returns {GoogleFormSubmitEvents} */
  static formSubmitEvents(v: any): GoogleFormSubmitEvents {
    return v as GoogleFormSubmitEvents;
  }
  /** @returns {GoogleCalenderUpdateEvents} */
  static calenderUpdateEvents(v: any): GoogleCalenderUpdateEvents {
    return v as GoogleCalenderUpdateEvents;
  }
  /** @returns { GoogleCalenderTimeDrivenEvents} */
  static calenderTImeDrivenEvents(v: any): GoogleCalenderTimeDrivenEvents {
    return v as GoogleCalenderTimeDrivenEvents;
  }
}
