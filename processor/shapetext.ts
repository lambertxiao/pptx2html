import PPTXProvider from '../provider';
import { Border, ShapeNode, SingleSlide, TextNode } from '../model';
import { computePixel, extractTextByPath } from '../util';
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
  globalCssStype: any

  constructor(provider: PPTXProvider, slide: SingleSlide, node: any, globalCssStype: any, withConnection: boolean) {
    super(provider, slide, node, globalCssStype)

    if (withConnection) {
      this.id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
      this.name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
      this.order = node["attrs"]["order"];
    } else {
      this.globalCssStype = globalCssStype
      this.id = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["id"];
      this.name = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["name"];

      let idx = (node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
      let type = (node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];

      this.order = node["attrs"]["order"];
      let slideLayoutSpNode = undefined;
      let slideMasterSpNode = undefined;

      if (type !== undefined) {
        if (idx !== undefined) {
          slideLayoutSpNode = this.slide.layoutIndexTables!["typeTable"][type];
          slideMasterSpNode = this.slide.masterIndexTable!["typeTable"][type];
        } else {
          slideLayoutSpNode = this.slide.layoutIndexTables!["typeTable"][type];
          slideMasterSpNode = this.slide.masterIndexTable!["typeTable"][type];
        }
      } else {
        if (idx !== undefined) {
          slideLayoutSpNode = this.slide.layoutIndexTables!["idxTable"][idx];
          slideMasterSpNode = this.slide.masterIndexTable!["idxTable"][idx];
        }
      }

      if (type === undefined) {
        type = extractTextByPath(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
        if (type === undefined) {
          type = extractTextByPath(slideMasterSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
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
      let { top, left } = this.getPosition(slideXfrmNode, undefined, undefined)
      let { width, height } = this.getSize(slideXfrmNode, undefined, undefined)

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

      // let headEndNodeAttrs = extractTextByPath(this.node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
      // let tailEndNodeAttrs = extractTextByPath(this.node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
      
      // // type: none, triangle, stealth, diamond, oval, arrow
      // if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
      //   (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
      //   let triangleMarker = "<defs><marker id=\"markerTriangle\" viewBox=\"0 0 10 10\" refX=\"1\" refY=\"5\" markerWidth=\"5\" markerHeight=\"5\" orient=\"auto-start-reverse\" markerUnits=\"strokeWidth\"><path d=\"M 0 0 L 10 5 L 0 10 z\" /></marker></defs>";
      //   result += triangleMarker;
      // }

      // switch (shapeType) {
      //   case "accentBorderCallout1":
      //   case "accentBorderCallout2":
      //   case "accentBorderCallout3":
      //   case "accentCallout1":
      //   case "accentCallout2":
      //   case "accentCallout3":
      //   case "actionButtonBackPrevious":
      //   case "actionButtonBeginning":
      //   case "actionButtonBlank":
      //   case "actionButtonDocument":
      //   case "actionButtonEnd":
      //   case "actionButtonForwardNext":
      //   case "actionButtonHelp":
      //   case "actionButtonHome":
      //   case "actionButtonInformation":
      //   case "actionButtonMovie":
      //   case "actionButtonReturn":
      //   case "actionButtonSound":
      //   case "arc":
      //   case "bevel":
      //   case "blockArc":
      //   case "borderCallout1":
      //   case "borderCallout2":
      //   case "borderCallout3":
      //   case "bracePair":
      //   case "bracketPair":
      //   case "callout1":
      //   case "callout2":
      //   case "callout3":
      //   case "can":
      //   case "chartPlus":
      //   case "chartStar":
      //   case "chartX":
      //   case "chevron":
      //   case "chord":
      //   case "cloud":
      //   case "cloudCallout":
      //   case "corner":
      //   case "cornerTabs":
      //   case "cube":
      //   case "decagon":
      //   case "diagStripe":
      //   case "diamond":
      //   case "dodecagon":
      //   case "donut":
      //   case "doubleWave":
      //   case "downArrowCallout":
      //   case "ellipseRibbon":
      //   case "ellipseRibbon2":
      //   case "flowChartAlternateProcess":
      //   case "flowChartCollate":
      //   case "flowChartConnector":
      //   case "flowChartDecision":
      //   case "flowChartDelay":
      //   case "flowChartDisplay":
      //   case "flowChartDocument":
      //   case "flowChartExtract":
      //   case "flowChartInputOutput":
      //   case "flowChartInternalStorage":
      //   case "flowChartMagneticDisk":
      //   case "flowChartMagneticDrum":
      //   case "flowChartMagneticTape":
      //   case "flowChartManualInput":
      //   case "flowChartManualOperation":
      //   case "flowChartMerge":
      //   case "flowChartMultidocument":
      //   case "flowChartOfflineStorage":
      //   case "flowChartOffpageConnector":
      //   case "flowChartOnlineStorage":
      //   case "flowChartOr":
      //   case "flowChartPredefinedProcess":
      //   case "flowChartPreparation":
      //   case "flowChartProcess":
      //   case "flowChartPunchedCard":
      //   case "flowChartPunchedTape":
      //   case "flowChartSort":
      //   case "flowChartSummingJunction":
      //   case "flowChartTerminator":
      //   case "folderCorner":
      //   case "frame":
      //   case "funnel":
      //   case "gear6":
      //   case "gear9":
      //   case "halfFrame":
      //   case "heart":
      //   case "heptagon":
      //   case "hexagon":
      //   case "homePlate":
      //   case "horizontalScroll":
      //   case "irregularSeal1":
      //   case "irregularSeal2":
      //   case "leftArrow":
      //   case "leftArrowCallout":
      //   case "leftBrace":
      //   case "leftBracket":
      //   case "leftRightArrowCallout":
      //   case "leftRightRibbon":
      //   case "irregularSeal1":
      //   case "lightningBolt":
      //   case "lineInv":
      //   case "mathDivide":
      //   case "mathEqual":
      //   case "mathMinus":
      //   case "mathMultiply":
      //   case "mathNotEqual":
      //   case "mathPlus":
      //   case "moon":
      //   case "nonIsoscelesTrapezoid":
      //   case "noSmoking":
      //   case "octagon":
      //   case "parallelogram":
      //   case "pentagon":
      //   case "pie":
      //   case "pieWedge":
      //   case "plaque":
      //   case "plaqueTabs":
      //   case "plus":
      //   case "quadArrowCallout":
      //   case "ribbon":
      //   case "ribbon2":
      //   case "rightArrowCallout":
      //   case "rightBrace":
      //   case "rightBracket":
      //   case "round1Rect":
      //   case "round2DiagRect":
      //   case "round2SameRect":
      //   case "rtTriangle":
      //   case "smileyFace":
      //   case "snip1Rect":
      //   case "snip2DiagRect":
      //   case "snip2SameRect":
      //   case "snipRoundRect":
      //   case "squareTabs":
      //   case "star10":
      //   case "star12":
      //   case "star16":
      //   case "star24":
      //   case "star32":
      //   case "star4":
      //   case "star5":
      //   case "star6":
      //   case "star7":
      //   case "star8":
      //   case "sun":
      //   case "teardrop":
      //   case "trapezoid":
      //   case "upArrowCallout":
      //   case "upDownArrowCallout":
      //   case "verticalScroll":
      //   case "wave":
      //   case "wedgeEllipseCallout":
      //   case "wedgeRectCallout":
      //   case "wedgeRoundRectCallout":
      //   case "rect":
      //     result += 
      //     `<rect x=0 y=0 width=${w} height=${h} fill=${fillColor}
      //       stroke=${border.color} stroke-width=${border.width} stroke-dasharray=${border.strokeDasharray} />`
      //     break;
      //   case "ellipse":
      //     result += "<ellipse cx='" + (w / 2) + "' cy='" + (h / 2) + "' rx='" + (w / 2) + "' ry='" + (h / 2) + "' fill='" + fillColor +
      //       "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
      //     break;
      //   case "roundRect":
      //     result += "<rect x='0' y='0' width='" + w + "' height='" + h + "' rx='7' ry='7' fill='" + fillColor +
      //       "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
      //     break;
      //   case "bentConnector2":    // 直角 (path)
      //     let d = "";
      //     if (isFlipV) {
      //       d = "M 0 " + w + " L " + h + " " + w + " L " + h + " 0";
      //     } else {
      //       d = "M " + w + " 0 L " + w + " " + h + " L 0 " + h;
      //     }
      //     result += "<path d='" + d + "' stroke='" + border.color +
      //       "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' fill='none' ";
      //     if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
      //       result += "marker-start='url(#markerTriangle)' ";
      //     }
      //     if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
      //       result += "marker-end='url(#markerTriangle)' ";
      //     }
      //     result += "/>";
      //     break;
      //   case "line":
      //   case "straightConnector1":
      //   case "bentConnector3":
      //   case "bentConnector4":
      //   case "bentConnector5":
      //   case "curvedConnector2":
      //   case "curvedConnector3":
      //   case "curvedConnector4":
      //   case "curvedConnector5":
      //     if (isFlipV) {
      //       result += "<line x1='" + w + "' y1='0' x2='0' y2='" + h + "' stroke='" + border.color +
      //         "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
      //     } else {
      //       result += "<line x1='0' y1='0' x2='" + w + "' y2='" + h + "' stroke='" + border.color +
      //         "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
      //     }
      //     if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
      //       result += "marker-start='url(#markerTriangle)' ";
      //     }
      //     if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
      //       result += "marker-end='url(#markerTriangle)' ";
      //     }
      //     result += "/>";
      //     break;
      //   case "rightArrow":
      //     result += "<defs><marker id=\"markerTriangle\" viewBox=\"0 0 10 10\" refX=\"1\" refY=\"5\" markerWidth=\"2.5\" markerHeight=\"2.5\" orient=\"auto-start-reverse\" markerUnits=\"strokeWidth\"><path d=\"M 0 0 L 10 5 L 0 10 z\" /></marker></defs>";
      //     result += "<line x1='0' y1='" + (h / 2) + "' x2='" + (w - 15) + "' y2='" + (h / 2) + "' stroke='" + border.color +
      //       "' stroke-width='" + (h / 2) + "' stroke-dasharray='" + border.strokeDasharray + "' ";
      //     result += "marker-end='url(#markerTriangle)' />";
      //     break;
      //   case "downArrow":
      //     result += "<defs><marker id=\"markerTriangle\" viewBox=\"0 0 10 10\" refX=\"1\" refY=\"5\" markerWidth=\"2.5\" markerHeight=\"2.5\" orient=\"auto-start-reverse\" markerUnits=\"strokeWidth\"><path d=\"M 0 0 L 10 5 L 0 10 z\" /></marker></defs>";
      //     result += "<line x1='" + (w / 2) + "' y1='0' x2='" + (w / 2) + "' y2='" + (h - 15) + "' stroke='" + border.color +
      //       "' stroke-width='" + (w / 2) + "' stroke-dasharray='" + border.strokeDasharray + "' ";
      //     result += "marker-end='url(#markerTriangle)' />";
      //     break;
      //   case "bentArrow":
      //   case "bentUpArrow":
      //   case "stripedRightArrow":
      //   case "quadArrow":
      //   case "circularArrow":
      //   case "swooshArrow":
      //   case "leftRightArrow":
      //   case "leftRightUpArrow":
      //   case "leftUpArrow":
      //   case "leftCircularArrow":
      //   case "notchedRightArrow":
      //   case "curvedDownArrow":
      //   case "curvedLeftArrow":
      //   case "curvedRightArrow":
      //   case "curvedUpArrow":
      //   case "upDownArrow":
      //   case "upArrow":
      //   case "uturnArrow":
      //   case "leftRightCircularArrow":
      //     break;
      //   case "triangle":
      //     break;
      //   case undefined:
      //   default:
      //     console.warn("Undefine shape type.");
      // }

      // result += "</svg>";

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

      // result += "<div class='block content " + this.getVerticalAlign() +
      //   "' _id='" + this.id + "' _idx='" + this.idx + "' _type='" + this.type + "' _name='" + this.name +
      //   this.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
      //   this.getBorder(false) +
      //   this.getShapeFill(false) +
      //   " z-index: " + this.order + ";" +
      //   "'>";

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
      return "initial"
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

    if (borderColor === undefined) {
      borderColor = "#000";
    } else {
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
