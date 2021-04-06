/**
 * Spreadsheetを操作する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
export function SpreadsheetLib(ss: GoogleAppsScript.Spreadsheet.Spreadsheet) {
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

/** 数値を3桁区切りの文字列に変換する */
export const to3DigitNum = (num: number): string => {
  if (num.toString() === 'NaN') return '0';
  return Number(num).toLocaleString();
};

/** 指定された桁数でゼロパディングする */
export const toZeroPadding = (v: string | number, digit: number): string => {
  const _ = digit < 0 ? -digit : digit; // 絶対値
  if (typeof v === 'number') {
    return `${'0'.repeat(_)}${v.toString()}`.slice(-_);
  } else {
    return `${'0'.repeat(_)}${v}`.slice(-_);
  }
};

/** 指定された文字列でパディングする */
export const toWordPadding = (v: string | number, digit: number, word: string): string => {
  const _ = digit < 0 ? -digit : digit; // 絶対値
  if (typeof v === 'number') {
    return `${word.repeat(_)}${v.toString()}`.slice(-_);
  } else {
    return `${word.repeat(_)}${v}`.slice(-_);
  }
};

/** 与えられた値に対して連続的に処理を行う */
export class Pipe<T> {
  private v: T;

  constructor(v: T) {
    this.v = v;
  }

  /** 途中経過をログ出力する */
  public log(): this {
    console.log(this.v);
    return this;
  }

  /** Pipe処理を継続する */
  public to(fc: (v: T) => T): this {
    this.v = fc(this.v);
    return this;
  }

  /** Pipe処理を完了する */
  public exit(): T {
    return this.v;
  }
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

class ArrayActions {
  /**
   * 指定されたIndexを１つ上に移動する
   */
  static swapUp<T>(array: T[], index: number): T[] {
    if (index <= 0) {
      return array;
    } else {
      array.splice(index - 1, 2, array[index], array[index - 1]);
      return array;
    }
  }

  /**
   * 指定されたIndexを１つ下に移動する
   */
  static swapDown<T>(array: T[], index: number): T[] {
    if (index < 0) {
      return array;
    } else if (array.length - 1 <= index) {
      return array;
    } else {
      console.log(array.length);
      array.splice(index, 2, array[index + 1], array[index]);
      return array;
    }
  }
}

export class TypeGuard {
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
