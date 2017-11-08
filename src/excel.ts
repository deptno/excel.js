import * as fs from 'fs'
import {Workbook} from 'excel4node'
import {Sheet} from './sheet'

export class Excel<T> {
  protected sheets: Sheet[] = []
  protected images: Set<string> = new Set()
  private wb

  constructor(sheetNames: string[], sheetProps?: T) {
    this.wb = new Workbook(defaultWorkbookOption)
    this.sheets = sheetNames.map((name, index) =>
      this.createSheet(this.wb.addWorksheet(name, defaultSheetOption), index, {
        name,
        ...sheetProps as any
      })
    )
  }

  public async render() {
    await Promise.all(this.sheets.map(sheet => sheet.render()))
    return this.buffer()
  }
 
  async buffer(): Promise<Buffer> {
    const buffer = await this.wb.writeToBuffer()
    this.removeImages()
    return buffer
  }

  protected createSheet(sheet, index: number, sheetProps: {name: string} & T): Sheet {
    throw new Error('return instance of YourSheet extends Sheet, `return new Sheet(sheet)`')
  }

  private removeImages() {
    process.nextTick(_ =>
      this.sheets.forEach(sheet =>
        sheet
          .getImages()
          .forEach(image => fs.unlink(image, err => err && console.error(err)) )
      )
    )
  }
}

const defaultWorkbookOption = {
  defaultFont : {
    jszip: {
      compression: 'DEFLATE'
    },
    name : '돋움',
    size : 9
  },
  numberFormat: '#,##0; -'
}
const defaultSheetOption = {
  sheetFormat: {
    defaultRowHeight: 17
  }
}
