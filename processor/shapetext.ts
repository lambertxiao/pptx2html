import PPTXProvider from '../provider';
import { Border, ShapeNode, SingleSlide } from '../model';
import { computePixel, extractTextByPath, printObj } from '../util';
import NodeProcessor from './processor';
const colz = require('colz');

export default class ShapeTextProcessor extends NodeProcessor {
  node: any
  slideLayoutSpNode: any
  slideMasterSpNode: any
  id: string
  name: string
  idx?: string
  type?: string
  order: string

  constructor(provider: PPTXProvider, slide: SingleSlide, node: any, withConnection: boolean) {
    super(provider, slide, node)

    if (withConnection) {
      this.id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
      this.name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
      this.order = node["attrs"]["order"];
    } else {
      this.id = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["id"];
      this.name = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["name"];

      let idx = (node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
      let type = (node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];

      this.order = node["attrs"]["order"];

      if (type) {
        this.slideLayoutSpNode = this.slide.layoutIndexTables!["typeTable"][type];
        this.slideMasterSpNode = this.slide.masterIndexTable!["typeTable"][type];
      } else {
        if (idx) {
          this.slideLayoutSpNode = this.slide.layoutIndexTables!["idxTable"][idx];
          this.slideMasterSpNode = this.slide.masterIndexTable!["idxTable"][idx];
        }
      }

      if (type === undefined) {
        type = extractTextByPath(this.slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
        if (type === undefined) {
          type = extractTextByPath(this.slideMasterSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
        }
      }

      this.type = type
      this.idx = idx
    }
  }

  async genHTML() {
    let node = this.node
    let xfrmList = ["p:spPr", "a:xfrm"];
    let slideXfrmNode = extractTextByPath(this.node, xfrmList);
    let slideLayoutXfrmNode = extractTextByPath(this.slideLayoutSpNode, xfrmList);
    let slideMasterXfrmNode = extractTextByPath(this.slideMasterSpNode, xfrmList);
    let shapeType = extractTextByPath(this.node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);

    let isFlipV = false;
    if (extractTextByPath(slideXfrmNode, ["attrs", "flipV"]) === "1" || extractTextByPath(slideXfrmNode, ["attrs", "flipH"]) === "1") {
      isFlipV = true;
    }

    let shapeNode: ShapeNode = {
      eleType: "shape",
      shapeType: shapeType,
      isFlipV: isFlipV,
    }

    if (shapeType) {
      let ext = extractTextByPath(slideXfrmNode, ["a:ext", "attrs"]);
      let w = computePixel(ext["cx"])
      let h = computePixel(ext["cy"])
      let { top, left } = this.getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
      let { width, height } = this.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)

      shapeNode.width = width
      shapeNode.height = height
      shapeNode.top = top
      shapeNode.left = left
      shapeNode.zindex = this.order
      shapeNode.ShapeWidth = w
      shapeNode.ShapeHeight = h

      // Fill Color
      let fillColor = this.getShapeFill()
      shapeNode.bgColor = fillColor

      // Border Color
      let border = this.getBorder()
      shapeNode.border = border

      // TextBody
      if (this.node["p:txBody"]) {
        let textNode = this.genTextBody(this.node["p:txBody"], this.type)

        if (textNode) {
          let tPosition = this.getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
          textNode.top = tPosition["top"]
          textNode.left = tPosition["left"]

          let tSize = this.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
          textNode.width = tSize["width"]
          textNode.height = tSize["height"]
          textNode.zindex = this.order

          shapeNode.textNode = textNode
        }
      }

      return shapeNode
    } else {
      let { top, left } = this.getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
      let { width, height } = this.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
      let border = this.getBorder()
      let bgColor = this.getShapeFill()

      // TextBody
      let textNode = this.genTextBody(node["p:txBody"], this.type);
      let sn: ShapeNode = {
        eleType: "shape",
        top: top,
        left: left,
        width: width,
        height: height,
        zindex: this.order,
        shapeType: shapeType,
        bgColor: bgColor,
        textNode: textNode,
        border: border,
      }

      return sn
    }
  }

  getShapeFill() {
    let node = this.node
    // 1. presentationML
    // p:spPr [a:noFill, solidFill, gradFill, blipFill, pattFill, grpFill]
    // From slide
    if (extractTextByPath(node, ["p:spPr", "a:noFill"]) !== undefined) {
      return "none"
    }

    let fillColor =  extractTextByPath(node, ["p:spPr", "a:solidFill", "a:srgbClr", "attrs", "val"]);
    // From theme
    if (fillColor === undefined) {
      let schemeClr = "a:" + extractTextByPath(node, ["p:spPr", "a:solidFill", "a:schemeClr", "attrs", "val"]);
      fillColor = this.getSchemeColor(schemeClr);
    }

    // 2. drawingML namespace
    if (fillColor === undefined) {
      let schemeClr = "a:" + extractTextByPath(node, ["p:style", "a:fillRef", "a:schemeClr", "attrs", "val"]);
      fillColor = this.getSchemeColor(schemeClr);
    }

    if (fillColor !== undefined) {
      fillColor = "#" + fillColor;

      // Apply shade or tint
      // TODO: 較淺, 較深 80%
      let lumMod = parseInt(extractTextByPath(node, ["p:spPr", "a:solidFill", "a:schemeClr", "a:lumMod", "attrs", "val"])) / 100000;
      let lumOff = parseInt(extractTextByPath(node, ["p:spPr", "a:solidFill", "a:schemeClr", "a:lumOff", "attrs", "val"])) / 100000;
      if (isNaN(lumMod)) {
        lumMod = 1.0;
      }
      if (isNaN(lumOff)) {
        lumOff = 0;
      }
      fillColor = this.applyLumModify(fillColor, lumMod, lumOff);
      return fillColor;
    } else {
      return fillColor;
    }
  }

  applyLumModify(rgbStr: string, factor: number, offset: number) {
    var color = new colz.Color(rgbStr);
    color.setLum(color.hsl.l * (1 + offset));
    return color.rgb.toString();
}

  getVerticalAlign() {
    // 上中下對齊: X, <a:bodyPr anchor="ctr">, <a:bodyPr anchor="b">
    let anchor = extractTextByPath(this.node, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
    if (anchor === undefined) {
      anchor = extractTextByPath(this.slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
      if (anchor === undefined) {
        anchor = extractTextByPath(this.slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
      }
    }

    return anchor === "ctr" ? "v-mid" : anchor === "b" ? "v-down" : "v-up";
  }

  getBorder(): Border {
    let node = this.node

    // 1. presentationML
    let lineNode = node["p:spPr"]["a:ln"];
    let borderWidthUnit = "pt"

    // Border width: 1pt = 12700, default = 0.75pt
    let borderWidth = parseInt(extractTextByPath(lineNode, ["attrs", "w"])) / 12700 / 5;
    if (isNaN(borderWidth) || borderWidth < 1) {
      borderWidth = 1
    }

    // Border color
    let borderColor = extractTextByPath(lineNode, ["a:solidFill", "a:srgbClr", "attrs", "val"]);
    if (borderColor === undefined) {
      let schemeClrNode = extractTextByPath(lineNode, ["a:solidFill", "a:schemeClr"]);
      let schemeClr = "a:" + extractTextByPath(schemeClrNode, ["attrs", "val"]);
      borderColor = this.getSchemeColor(schemeClr);
    }

    // 2. drawingML namespace
    if (borderColor === undefined) {
      let schemeClrNode = extractTextByPath(node, ["p:style", "a:lnRef", "a:schemeClr"]);
      let schemeClr = "a:" + extractTextByPath(schemeClrNode, ["attrs", "val"]);
      let borderColor = this.getSchemeColor(schemeClr);

      if (borderColor !== undefined) {
        let shade = extractTextByPath(schemeClrNode, ["a:shade", "attrs", "val"]);
        if (shade !== undefined) {
          shade = parseInt(shade) / 100000;
          let color = new colz.Color("#" + borderColor);
          color.setLum(color.hsl.l * shade);
          borderColor = color.hex.replace("#", "");
        }
      }
    }

    if (borderColor) {
      borderColor = "#" + borderColor;
    }

    // Border type
    let _borderType = extractTextByPath(lineNode, ["a:prstDash", "attrs", "val"]);
    let borderType: string
    let strokeDasharray = "0";
    switch (_borderType) {
      case "solid":
        borderType = "solid";
        strokeDasharray = "0";
        break;
      case "dash":
        borderType = "dashed";
        strokeDasharray = "5";
        break;
      case "dashDot":
        borderType = "dashed";
        strokeDasharray = "5, 5, 1, 5";
        break;
      case "dot":
        borderType = "dotted";
        strokeDasharray = "1, 5";
        break;
      case "lgDash":
        borderType = "dashed";
        strokeDasharray = "10, 5";
        break;
      case "lgDashDotDot":
        borderType = "dashed";
        strokeDasharray = "10, 5, 1, 5, 1, 5";
        break;
      case "sysDash":
        borderType = "dashed";
        strokeDasharray = "5, 2";
        break;
      case "sysDashDot":
        borderType = "dashed";
        strokeDasharray = "5, 2, 1, 5";
        break;
      case "sysDashDotDot":
        borderType = "dashed";
        strokeDasharray = "5, 2, 1, 5, 1, 5";
        break;
      case "sysDot":
        borderType = "dotted";
        strokeDasharray = "2, 5";
        break;
      case undefined:
      default:
        borderType = "solid";
        strokeDasharray = "0";
        break;
    }

    let border: Border = {
      color: borderColor,
      width: borderWidth,
      widthUnit: borderWidthUnit,
      type: borderType,
      strokeDasharray: strokeDasharray,
    }

    return border
  }
}
