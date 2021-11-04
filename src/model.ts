import { extractText } from "./util"

export class GlobalProps {
  slideWidth?: number
  slideHeight?: number
  slidePaths?: any[]
  slideLayouts?: any[]
  thumbnail?: string
  theme?: ThemeContent
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

  getTargetFromSlide(rid: string) {
    if (this.resContent[rid]) {
      return this.resContent[rid]["target"];
    }

    return ""
  }

  getTargetFromLayout(rid: string): string {
   let lrss = this.layoutResContent["Relationships"]["Relationship"]
    for (const rs of lrss) {
      if (rs["attrs"]["Id"] == rid) {
        return rs["attrs"]["Target"].replace("../", "ppt/");
      }
    }

    return ""
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
    let refNode = extractText(this.content, ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
    let color = extractText(refNode, ["a:srgbClr", "attrs", "val"]);
    if (color === undefined) {
      color = extractText(refNode, ["a:sysClr", "attrs", "lastClr"]);
    }

    return color;
  }
}

export class NodeElement {
  name?: string
  eleType?: string
  zindex?: string
  width?: number
  height?: number
  top?: number
  left?: number
}

export interface Border {
  color: string
  type: string
  width: number
  widthUnit: string
  strokeDasharray: any
}

export class PicNode extends NodeElement {
  imgUrl?: string
  mimeType?: string
  borderRadius?: number
}

export class ShapeNode extends NodeElement {
  shapeType?: string
  bgColor?: string
  fontColor?: string
  bgImg?: string
  textNode?: TextNode
  border?: Border
  ShapeWidth?: number
  ShapeHeight?: number
  isFlipV?: boolean
}

export class TextNode extends NodeElement {
  eleType = "text"
  textType?: string
  color?: string
  fontSize?: number
  fontFamily?: string
  content?: SpanNode | null
  styleClass?: string
  paragraphNodes: ParagraphNode[] = []
}

export class ParagraphNode {
  spans: SpanNode[] = []
}

export class SpanNode {
  marginLeft?: number
  marginRight?: number
  fontSize?: number | string
  fontSizeUnit?: string
  fontFamily?: string
  fontStyle?: string
  textDecoration?: string
  verticalAlign?: string
  content?: string
  color?: string
  // 有超链接
  linkID?: string
}

export class TableCol {
  rowSpan?: number
  colSpan?: number
  text?: TextNode
}

export class TableRow {
  cols: TableCol[] = []
}

export class TableNode extends NodeElement {
  eleType = "table"
  rows: TableRow[] = []
}

export class ChartNode extends NodeElement {
  eleType = "chart"
}

export class DiagramNode extends NodeElement {
  eleType = "diagram"
}

export class NodeElementGroup extends NodeElement {
  eleType = "nodeGroup"
  nodes: NodeElement[] = []
}

export class SlideView {
  width?: number
  height?: number
  bgColor?: string
  bgImgData?: string
  layoutNodes: NodeElement[] = []
  slideNodes: NodeElement[] = []

  addLayoutNode(node: NodeElement) {
    this.layoutNodes.push(node)
  }

  addSlideNode(node: NodeElement) {
    this.slideNodes.push(node)
  }
}
