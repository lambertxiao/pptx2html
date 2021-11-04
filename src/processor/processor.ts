import PPTXProvider from '../provider';
import { NodeElement, ParagraphNode, SingleSlide, SpanNode, TextNode } from '../model';
import { computePixel, extractText, printObj } from '../util';

export default abstract class NodeProcessor {
  provider: PPTXProvider
  slide: SingleSlide
  node: any

  constructor(
    provider: PPTXProvider,
    slide: SingleSlide,
    node: any,
  ) {
    this.provider = provider
    this.slide = slide
    this.node = node
  }

  abstract process(): Promise<NodeElement | null>

  getSchemeColor(clr: string) {
    return this.slide.gprops!.theme!.getSchemeColor(clr)
  }

  getPosition(slideSpNode: any, slideLayoutSpNode: any, slideMasterSpNode: any) {
    let off = undefined;
    let x = -1, y = -1;

    if (slideSpNode) {
      off = slideSpNode["a:off"]["attrs"];
    } else if (slideLayoutSpNode) {
      off = slideLayoutSpNode["a:off"]["attrs"];
    } else if (slideMasterSpNode) {
      off = slideMasterSpNode["a:off"]["attrs"];
    }

    if (off === undefined) {
      return { top: 0, left: 0 };
    } else {
      x = computePixel(off["x"])
      y = computePixel(off["y"])
      return { top: y, left: x }
    }
  }

  getSize(slideSpNode: any, slideLayoutSpNode: any, slideMasterSpNode: any) {
    let ext = undefined;
    let w = -1, h = -1;

    if (slideSpNode) {
      ext = slideSpNode["a:ext"]["attrs"];
    } else if (slideLayoutSpNode) {
      ext = slideLayoutSpNode["a:ext"]["attrs"];
    } else if (slideMasterSpNode) {
      ext = slideMasterSpNode["a:ext"]["attrs"];
    }

    if (ext === undefined) {
      return { width: 0, height: 0 }
    } else {
      w = computePixel(ext["cx"])
      h = computePixel(ext["cy"])
      return { width: w, height: h }
    }
  }

  genBuChar(node: any): SpanNode | null {
    let pPrNode = node["a:pPr"];
    let lvl = parseInt(extractText(pPrNode, ["attrs", "lvl"]));
    if (isNaN(lvl)) {
      lvl = 0;
    }

    let buChar = extractText(pPrNode, ["a:buChar", "attrs", "char"]);
    if (!buChar) {
      return null
    }

    let buFontAttrs = extractText(pPrNode, ["a:buFont", "attrs"]);
    let spanNode: SpanNode = {
      content: buChar,
    }

    if (buFontAttrs) {
      let marginLeft = parseInt(extractText(pPrNode, ["attrs", "marL"])) * 96 / 914400;
      let marginRight = parseInt(buFontAttrs["pitchFamily"]);

      if (isNaN(marginLeft)) {
        marginLeft = 328600 * 96 / 914400;
      }
      if (isNaN(marginRight)) {
        marginRight = 0;
      }

      spanNode.fontFamily = buFontAttrs["typeface"]
      spanNode.marginLeft = marginLeft * lvl
      spanNode.marginRight = marginRight
      spanNode.fontSize = 20
      spanNode.fontSizeUnit = "pt"

      return spanNode
    } else {
      let marginLeft = 328600 * 96 / 914400 * lvl;
      spanNode.marginLeft = marginLeft
      return spanNode
    }
  }

  genTextBody(textBodyNode: any, type?: string): TextNode | undefined {
    if (!textBodyNode) {
      return undefined;
    }

    let textNode = new TextNode()
    textNode.eleType = "text"

    if (textBodyNode["a:p"].constructor === Array) {
      // 多个文本段
      for (let i = 0; i < textBodyNode["a:p"].length; i++) {
        let ph = new ParagraphNode()
        let pNode = textBodyNode["a:p"][i];
        let rNode = pNode["a:r"];
        textNode.styleClass = this.getHorizontalAlign(pNode, type)
        textNode.content = this.genBuChar(pNode)

        if (!rNode) {
          ph.spans.push(this.genSpanElement(pNode, type))
        } else if (rNode.constructor === Array) {
          for (let j = 0; j < rNode.length; j++) {
            ph.spans.push(this.genSpanElement(rNode[j], type))
          }
        } else {
          ph.spans.push(this.genSpanElement(rNode, type))
        }

        textNode.paragraphNodes.push(ph)
      }
    } else {
      let ph = new ParagraphNode()
      // 单个文本段
      let pNode = textBodyNode["a:p"]
      let rNode = pNode["a:r"]
      let styleClass = this.getHorizontalAlign(pNode, type)

      textNode.styleClass = styleClass
      let content = this.genBuChar(pNode)
      if (content) {
        textNode.content = content
      }

      if (!rNode) {
        ph.spans.push(this.genSpanElement(pNode, type))
      } else if (rNode.constructor === Array) {
        for (let j = 0; j < rNode.length; j++) {
          ph.spans.push(this.genSpanElement(rNode[j], type))
        }
      } else {
        ph.spans.push(this.genSpanElement(rNode, type))
      }

      textNode.paragraphNodes.push(ph)
    }

    return textNode;
  }

  getHorizontalAlign(node: any, type?: string) {
    let algn = extractText(node, ["a:pPr", "attrs", "algn"]);
    if (algn === undefined) {
      algn = extractText(this.slide.layoutContent, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
      if (algn === undefined) {
        algn = extractText(this.slide.masterContent, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
        if (algn === undefined) {
          switch (type) {
            case "title":
            case "subTitle":
            case "ctrTitle":
              algn = extractText(this.slide.masterTextStyles, ["p:titleStyle", "a:lvl1pPr", "attrs", "alng"]);
              break;
            default:
              algn = extractText(this.slide.masterTextStyles, ["p:otherStyle", "a:lvl1pPr", "attrs", "alng"]);
          }
        }
      }
    }
    // TODO:
    if (algn === undefined) {
      if (type == "title" || type == "subTitle" || type == "ctrTitle") {
        return "h-mid";
      } else if (type == "sldNum") {
        return "h-right";
      }
    }
    return algn === "ctr" ? "h-mid" : algn === "r" ? "h-right" : "h-left";
  }

  genSpanElement(node: any, type?: string): SpanNode {
    let text = node["a:t"];
    if (typeof text !== 'string') {
      text = extractText(node, ["a:fld", "a:t"]);
      if (typeof text !== 'string') {
        text = "";
      }
    }

    let sn: SpanNode = {
      color: this.getFontColor(node),
      fontSize: this.getFontSize(node, this.slide.layoutResContent, type, this.slide.masterTextStyles!),
      fontFamily: this.getFontType(node, type),
      fontStyle: this.getFontItalic(node),
      textDecoration: this.getFontDecoration(node),
      verticalAlign: this.getTextVerticalAlign(node),
      linkID: extractText(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]),
      content: text,
    }

    return sn
  }

  getFontType(node: any, type: any) {
    let typeface = extractText(node, ["a:rPr", "a:latin", "attrs", "typeface"]);

    if (typeface === undefined) {
      let fontSchemeNode = extractText(this.slide!.gprops!.theme, ["a:theme", "a:themeElements", "a:fontScheme"]);
      if (type == "title" || type == "subTitle" || type == "ctrTitle") {
        typeface = extractText(fontSchemeNode, ["a:majorFont", "a:latin", "attrs", "typeface"]);
      } else if (type == "body") {
        typeface = extractText(fontSchemeNode, ["a:minorFont", "a:latin", "attrs", "typeface"]);
      } else {
        typeface = extractText(fontSchemeNode, ["a:minorFont", "a:latin", "attrs", "typeface"]);
      }
    }

    return (typeface === undefined) ? "inherit" : typeface;
  }

  getTextByPathStr(node: any, pathStr: any) {
    return extractText(node, pathStr.trim().split(/\s+/));
  }

  getFontColor(node: any) {
    let color = this.getTextByPathStr(node, "a:rPr a:solidFill a:srgbClr attrs val");
    return (color === undefined) ? "" : "#" + color;
  }

  getFontSize(node: any, slideLayoutSpNode: any, type: any, slideMasterTextStyles: any) {
    let fontSize: any;
    if (node["a:rPr"]) {
      fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
    }

    if ((isNaN(fontSize) || fontSize === undefined)) {
      let sz = extractText(slideLayoutSpNode, ["p:txBody", "a:lstStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
      fontSize = parseInt(sz) / 100;
    }

    if (isNaN(fontSize) || fontSize === undefined) {
      let sz: any
      if (type == "title" || type == "subTitle" || type == "ctrTitle") {
        sz = extractText(slideMasterTextStyles, ["p:titleStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
      } else if (type == "body") {
        sz = extractText(slideMasterTextStyles, ["p:bodyStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
      } else if (type == "dt" || type == "sldNum") {
        sz = "1200";
      } else if (type === undefined) {
        sz = extractText(slideMasterTextStyles, ["p:otherStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
      }

      fontSize = parseInt(sz) / 100;
    }

    let baseline = extractText(node, ["a:rPr", "attrs", "baseline"]);
    if (baseline && !isNaN(fontSize)) {
      fontSize -= 10;
    }

    return isNaN(fontSize) ? "inherit" : (fontSize + "pt");
  }

  getFontBold(node: any) {
    return (node["a:rPr"] && node["a:rPr"]["attrs"]["b"] === "1") ? "bold" : "initial";
  }

  getFontItalic(node: any) {
    return (node["a:rPr"] && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "normal";
  }

  getFontDecoration(node: any) {
    return (node["a:rPr"] && node["a:rPr"]["attrs"]["u"] === "sng") ? "underline" : "initial";
  }

  getTextVerticalAlign(node: any) {
    let baseline = extractText(node, ["a:rPr", "attrs", "baseline"]);
    return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
  }
}