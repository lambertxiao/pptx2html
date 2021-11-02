import PPTXProvider from '../provider';
import { NodeElement, ParagraphNode, SingleSlide, SpanNode, TextNode } from '../model';
import { computePixel, extractTextByPath } from '../util';

export default abstract class NodeProcessor {
  provider: PPTXProvider
  slide: SingleSlide
  node: any
  globalCssStyles: any

  constructor(
    provider: PPTXProvider,
    slide: SingleSlide,
    node: any,
    globalCssStyles: any
  ) {
    this.provider = provider
    this.slide = slide
    this.node = node
    this.globalCssStyles = globalCssStyles
  }

  abstract genHTML(): Promise<NodeElement | null>

  getSchemeColor(clr: string) {
    return this.slide.gprops!.theme!.getSchemeColor(clr)
  }

  getPosition(slideSpNode: any, slideLayoutSpNode: any, slideMasterSpNode: any) {
    let off = undefined;
    let x = -1, y = -1;

    if (slideSpNode !== undefined) {
      off = slideSpNode["a:off"]["attrs"];
    } else if (slideLayoutSpNode !== undefined) {
      off = slideLayoutSpNode["a:off"]["attrs"];
    } else if (slideMasterSpNode !== undefined) {
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

    if (slideSpNode !== undefined) {
      ext = slideSpNode["a:ext"]["attrs"];
    } else if (slideLayoutSpNode !== undefined) {
      ext = slideLayoutSpNode["a:ext"]["attrs"];
    } else if (slideMasterSpNode !== undefined) {
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
    let lvl = parseInt(extractTextByPath(pPrNode, ["attrs", "lvl"]));
    if (isNaN(lvl)) {
      lvl = 0;
    }

    let buChar = extractTextByPath(pPrNode, ["a:buChar", "attrs", "char"]);
    if (!buChar) {
      return null
    }

    let buFontAttrs = extractTextByPath(pPrNode, ["a:buFont", "attrs"]);
    let spanNode: SpanNode = {
      content: buChar,
    }

    if (buFontAttrs !== undefined) {
      let marginLeft = parseInt(extractTextByPath(pPrNode, ["attrs", "marL"])) * 96 / 914400;
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
    let algn = extractTextByPath(node, ["a:pPr", "attrs", "algn"]);
    if (algn === undefined) {
      algn = extractTextByPath(this.slide.layoutContent, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
      if (algn === undefined) {
        algn = extractTextByPath(this.slide.masterContent, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
        if (algn === undefined) {
          switch (type) {
            case "title":
            case "subTitle":
            case "ctrTitle":
              algn = extractTextByPath(this.slide.masterTextStyles, ["p:titleStyle", "a:lvl1pPr", "attrs", "alng"]);
              break;
            default:
              algn = extractTextByPath(this.slide.masterTextStyles, ["p:otherStyle", "a:lvl1pPr", "attrs", "alng"]);
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
      text = extractTextByPath(node, ["a:fld", "a:t"]);
      if (typeof text !== 'string') {
        text = "&nbsp;";
      }
    }

    let sn: SpanNode = {
      color: this.getFontColor(node),
      fontSize: this.getFontSize(node, this.slide.layoutResContent, type, this.slide.masterTextStyles!),
      fontFamily: this.getFontType(node, type),
      fontStyle: this.getFontItalic(node),
      textDecoration: this.getFontDecoration(node),
      verticalAlign: this.getTextVerticalAlign(node),
      linkID: extractTextByPath(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]),
      content: text,
    }
    
    return sn
  }

  getFontType(node: any, type: any) {
    let typeface = extractTextByPath(node, ["a:rPr", "a:latin", "attrs", "typeface"]);

    if (typeface === undefined) {
      let fontSchemeNode = extractTextByPath(this.slide!.gprops!.theme, ["a:theme", "a:themeElements", "a:fontScheme"]);
      if (type == "title" || type == "subTitle" || type == "ctrTitle") {
        typeface = extractTextByPath(fontSchemeNode, ["a:majorFont", "a:latin", "attrs", "typeface"]);
      } else if (type == "body") {
        typeface = extractTextByPath(fontSchemeNode, ["a:minorFont", "a:latin", "attrs", "typeface"]);
      } else {
        typeface = extractTextByPath(fontSchemeNode, ["a:minorFont", "a:latin", "attrs", "typeface"]);
      }
    }

    return (typeface === undefined) ? "inherit" : typeface;
  }

  getTextByPathStr(node: any, pathStr: any) {
    return extractTextByPath(node, pathStr.trim().split(/\s+/));
  }

  getFontColor(node: any) {
    let color = this.getTextByPathStr(node, "a:rPr a:solidFill a:srgbClr attrs val");
    return (color === undefined) ? "" : "#" + color;
  }

  getFontSize(node: any, slideLayoutSpNode: any, type: any, slideMasterTextStyles: any) {
    let fontSize: any;
    if (node["a:rPr"] !== undefined) {
      fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
    }

    if ((isNaN(fontSize) || fontSize === undefined)) {
      let sz = extractTextByPath(slideLayoutSpNode, ["p:txBody", "a:lstStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
      fontSize = parseInt(sz) / 100;
    }

    if (isNaN(fontSize) || fontSize === undefined) {
      let sz: any
      if (type == "title" || type == "subTitle" || type == "ctrTitle") {
        sz = extractTextByPath(slideMasterTextStyles, ["p:titleStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
      } else if (type == "body") {
        sz = extractTextByPath(slideMasterTextStyles, ["p:bodyStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
      } else if (type == "dt" || type == "sldNum") {
        sz = "1200";
      } else if (type === undefined) {
        sz = extractTextByPath(slideMasterTextStyles, ["p:otherStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
      }

      fontSize = parseInt(sz) / 100;
    }

    let baseline = extractTextByPath(node, ["a:rPr", "attrs", "baseline"]);
    if (baseline !== undefined && !isNaN(fontSize)) {
      fontSize -= 10;
    }

    return isNaN(fontSize) ? "inherit" : (fontSize + "pt");
  }

  getFontBold(node: any) {
    return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["b"] === "1") ? "bold" : "initial";
  }

  getFontItalic(node: any) {
    return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "normal";
  }

  getFontDecoration(node: any) {
    return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["u"] === "sng") ? "underline" : "initial";
  }

  getTextVerticalAlign(node: any) {
    let baseline = extractTextByPath(node, ["a:rPr", "attrs", "baseline"]);
    return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
  }
}