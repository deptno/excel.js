export abstract class Sheet {
  protected images: Set<string> = new Set()

  constructor(protected sheet) {
  }

  public async render() {
    this.setupSheet()
    return await Promise.all([
      this.renderLayout(),
      this.renderBody()
    ])
  }

  protected setupSheet() {
    throw new Error('must do implement')
  }

  protected renderLayout() {
    throw new Error('must do implement')
  }

  protected renderBody() { throw new Error('must do implement')
  }

  public getImages() {
    return this.images
  }

  protected mergeCell(posLt: Coordinate, posRb: Coordinate): this {
    this.sheet.cell(posLt[0], posLt[1], posRb[0], posRb[1], true)
    return this
  }

  protected mergeCellWithDraw(posLt: Coordinate, posRb: Coordinate, content, style?): this {
    const cell = this.sheet.cell(posLt[0], posLt[1], posRb[0], posRb[1], true).string(content)

    if (style) {
      cell.style(style)
    }
    return this
  }

  protected setRowHeight(row: number, height: number): this {
    this.sheet.row(row).setHeight(height)
    return this
  }

  protected setColumnWidth(column: number, width: number): this {
    this.sheet.column(column).setWidth(width)
    return this
  }

  protected draw([x, y]: Coordinate, content, style?, type?: 'formula'|'number'|'string', link?): this {
    const cell = this.sheet.cell(x, y)

    if (type === 'formula') {
      cell.formula(content)
    } else if (type === 'number') {
      cell.number(content)
    } else {
      cell.string(content)
    }
    if (link) {
      const {url, text, toolTip} = link
      cell.link(url, text, toolTip)
    }
    if (style) {
      cell.style(style)
    }
    return this
  }

  protected paint(posLt: Coordinate, posRb: Coordinate, color = 'white'): this {
    this.sheet
      .cell(posLt[0], posLt[1], posRb[0], posRb[1])
      .style({
        fill: {
          type       : 'pattern',
          patternType: 'solid',
          fgColor    : color
        }
      })
    return this
  }

  protected setAnchorImage([x, y]: Coordinate, path) {
    this.sheet.addImage({
      path,
      type    : 'picture',
      position: {
        type: 'absoluteAnchor',
        x   : `${x}mm`,
        y   : `${y}mm`
      }
    })
  }

  protected setTwoCellAnchorImage([fromCol, fromRow]: Coordinate, [toCol, toRow]: Coordinate, path) {
    this.sheet.addImage({
      path,
      type    : 'picture',
      position: {
        type: 'twoCellAnchor',
        from: {
          col: fromCol,
          row: fromRow
        },
        to  : {
          col: toCol,
          row: toRow
        }
      }
    })
  }

  protected get() {
    return this.sheet
  }
}

type Coordinate = [number, number]
