import { extractTextByPath } from "./util"

export class GlobalProps {
  slideWidth?: number
  slideHeight?: number
  slidePaths?: any[]
  slideLayouts?: any[]
  thumbnail?: string
  theme?: ThemeContent

  globalStyles: { [key: string]: CssStyle } = {}

  addStyle(name: string, val: CssStyle) {
    this.globalStyles[name] = val
  }
}

export class SingleSlide {
  content: any
  // slide的资源内容
  resContent: any
  // 布局文件的内容
  layoutContent: any
  // 布局文件的资源内容
  layoutResContent: any
  // 布局文件内容索引表
  layoutIndexTables?: { idTable: any, idxTable: any, typeTable: any }
  // 母版的内容
  masterContent: any
  // 母版的字体样式
  masterTextStyles?: string
  // 母版的索引表
  masterIndexTable?: { idTable: any, idxTable: any, typeTable: any }
  masterResContent?: any
  bgColor?: string
  gprops?: GlobalProps

  getTargetResByID(rid: string) {
    if (this.resContent[rid]) {
      return this.resContent[rid]["target"];
    }

    let lrss = this.layoutResContent["Relationships"]["Relationship"]
    for (const rs of lrss) {
      if (rs["attrs"]["Id"] == rid) {
        return rs["attrs"]["Target"].replace("../", "ppt/");
      }
    }

    return ""
  }

  queryLayoutIndex() { }
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
    let s = ""
    for (const key in this.content) {
      s += `${key}: ${this.content[key]};`
    }

    return `.${this.name} {${s}}`
  }
}

export class ThemeContent {

  constructor(private readonly content: any) { }

  getSchemeColor(schemeClr: string) {
    switch (schemeClr) {
      case "tx1": schemeClr = "a:dk1"; break;
      case "tx2": schemeClr = "a:dk2"; break;
      case "bg1": schemeClr = "a:lt1"; break;
      case "bg2": schemeClr = "a:lt2"; break;
    }
    let refNode = extractTextByPath(this.content, ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
    let color = extractTextByPath(refNode, ["a:srgbClr", "attrs", "val"]);
    if (color === undefined) {
      color = extractTextByPath(refNode, ["a:sysClr", "attrs", "lastClr"]);
    }

    return color;
  }
}

export class Node {

  constructor(private readonly content: any) { }

  getSubNode() {

  }
}
