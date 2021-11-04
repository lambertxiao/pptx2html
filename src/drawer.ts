import { SlideView } from "./model"

export interface Drawer {
  draw(nodes: SlideView[]): string
}

export class CssStyle {

  content: { [key: string]: string } = {}

  constructor(private readonly name: string) {
    this.name = name
  }

  add(key: string, val: string) {
    this.content[key] = val
  }

  addWidth(val: number) {
    this.add("width", val + "px")
  }

  addHeight(val: number) {
    this.add("height", val + "px")
  }

  addBGBase64Img(val: string) {
    this.add("background-image", `url(data:image/png;base64,${val})`)
  }

  toString() {
    let s = `.${this.name} {`
    for (const key in this.content) {
      s += `${key}: ${this.content[key]};`
    }

    s += "}"
    return s
  }
}

