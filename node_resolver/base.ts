import PPTXProvider from '../pptx_provider';
import { SingleSlide } from '../slide';
import { computePixel, extractTextByPath } from '../util';

export default abstract class NodeResolver {
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

	abstract genHTML(): Promise<string>

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
      return "";
    } else {
      x = computePixel(off["x"])
      y = computePixel(off["y"])
      return (isNaN(x) || isNaN(y)) ? "" : "top:" + y + "px; left:" + x + "px;";
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
      return "";
    } else {
      w = computePixel(ext["cx"])
      h = computePixel(ext["cy"])
      return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
    }
  }

	genBuChar(node: any) {
    let pPrNode = node["a:pPr"];
    let lvl = parseInt(extractTextByPath(pPrNode, ["attrs", "lvl"]));
    if (isNaN(lvl)) {
      lvl = 0;
    }

    let buChar = extractTextByPath(pPrNode, ["a:buChar", "attrs", "char"]);
    if (buChar !== undefined) {
      let buFontAttrs = extractTextByPath(pPrNode, ["a:buFont", "attrs"]);
      if (buFontAttrs !== undefined) {
        let marginLeft = parseInt(extractTextByPath(pPrNode, ["attrs", "marL"])) * 96 / 914400;
        let marginRight = parseInt(buFontAttrs["pitchFamily"]);
        
        if (isNaN(marginLeft)) {
          marginLeft = 328600 * 96 / 914400;
        }
        if (isNaN(marginRight)) {
          marginRight = 0;
        }
        
        let typeface = buFontAttrs["typeface"];

        return "<span style='font-family: " + typeface +
          "; margin-left: " + marginLeft * lvl + "px" +
          "; margin-right: " + marginRight + "px" +
          "; font-size: 20pt" +
          "'>" + buChar + "</span>";
      } else {
        let marginLeft = 328600 * 96 / 914400 * lvl;
        return "<span style='margin-left: " + marginLeft + "px;'>" + buChar + "</span>";
      }
    } else {
      //buChar = '•';
      return "<span style='margin-left: " + 328600 * 96 / 914400 * lvl + "px" +
        "; margin-right: " + 0 + "px;'></span>";
    }
  }

	genTextBody(textBodyNode: any, type?: string) {
    let text = "";
    if (textBodyNode === undefined) {
      return text;
    }

    if (textBodyNode["a:p"].constructor === Array) {
      // multi p
      for (let i = 0; i < textBodyNode["a:p"].length; i++) {
        let pNode = textBodyNode["a:p"][i];
        let rNode = pNode["a:r"];
        text += "<div class='" + this.getHorizontalAlign(pNode, type) + "'>";
        text += this.genBuChar(pNode);
        if (rNode === undefined) {
          // without r
          text += this.genSpanElement(pNode, type);
        } else if (rNode.constructor === Array) {
          // with multi r
          for (let j = 0; j < rNode.length; j++) {
            text += this.genSpanElement(rNode[j], type);
          }
        } else {
          // with one r
          text += this.genSpanElement(rNode, type);
        }
        text += "</div>";
      }
    } else {
      // one p
      let pNode = textBodyNode["a:p"];
      let rNode = pNode["a:r"];
      text += "<div class='" + this.getHorizontalAlign(pNode, type) + "'>";
      text += this.genBuChar(pNode);
      if (rNode === undefined) {
        // without r
        text += this.genSpanElement(pNode, type);
      } else if (rNode.constructor === Array) {
        // with multi r
        for (let j = 0; j < rNode.length; j++) {
          text += this.genSpanElement(rNode[j], type);
        }
      } else {
        // with one r
        text += this.genSpanElement(rNode, type);
      }
      text += "</div>";
    }

    return text;
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

	genSpanElement(node: any, type?: string) {
    let slideMasterTextStyles = this.slide.masterTextStyles
    let text = node["a:t"];
    if (typeof text !== 'string') {
      text = extractTextByPath(node, ["a:fld", "a:t"]);
      if (typeof text !== 'string') {
        text = "&nbsp;";
      }
    }

    let styleText =
      "color:" + this.getFontColor(node) +
      ";font-size:" + this.getFontSize(node, this.slide.layoutResContent, type, this.slide.masterTextStyles!) +
      ";font-family:" + this.getFontType(node, type) +
      ";font-weight:" + this.getFontBold(node) +
      ";font-style:" + this.getFontItalic(node) +
      ";text-decoration:" + this.getFontDecoration(node) +
      ";vertical-align:" + this.getTextVerticalAlign(node) +
      ";";

    let cssName = "";
    let globalCssStyles = this.globalCssStyles

    if (styleText in globalCssStyles) {
      cssName = globalCssStyles[styleText]["name"];
    } else {
      cssName = "_css_" + (Object.keys(globalCssStyles).length + 1);
      globalCssStyles[styleText] = {
        "name": cssName,
        "text": styleText
      };
    }

    let linkID = extractTextByPath(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
    if (linkID !== undefined) {
      let linkURL = this.slide.resContent[linkID]["target"];
      return "<span class='text-block " + cssName + "'><a href='" + linkURL + "' target='_blank'>" + text.replace(/\s/i, "&nbsp;") + "</a></span>";
    } else {
      return "<span class='text-block " + cssName + "'>" + text.replace(/\s/i, "&nbsp;") + "</span>";
    }
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
    return (color === undefined) ? "#000" : "#" + color;
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