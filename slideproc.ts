import PPTXProvider from './provider';
import { CssStyle, GlobalProps, NodeElement, SingleSlide } from './model';
import { extractTextByPath, getSchemeColorFromTheme, img2Base64 } from './util';
import PicProcessor from './processor/pic';
import ShapeTextProcessor from './processor/shapetext';
import GraphicProcessor from './processor/graphic'

export default class SlideProcessor {
  // slide的内容节点
  slideNodes?: any
  slideLayoutNodes?: any

  gprops: GlobalProps
  slide?: SingleSlide
  layoutBg: any
  masterBg: any

  constructor(
    private readonly slidePath: string,
    private readonly index: number,
    private readonly provider: PPTXProvider,
    gprops: GlobalProps,
    private readonly globalCssStyles: any,
  ) {
    this.gprops = gprops
  }

  async process() {
    await this.prepare()
    return await this.genHtml()
  }

  async prepare() {
    let slide = new SingleSlide()
    this.slide = slide
    this.slide.gprops = this.gprops
    slide.content = await this.provider.loadXML(this.slidePath)

    let { layoutFilePath, slideResContent } = await this.getSlideRes(this.slidePath)
    slide.resContent = slideResContent

    slide.layoutContent = await this.provider.loadXML(layoutFilePath)
    slide.layoutIndexTables = await this.indexNodes(slide.layoutContent)
    slide.layoutResContent = await this.getSlideLayoutRes(layoutFilePath)

    let masterFilePath = this.getSlideMasterFilePath(slide.layoutResContent)
    slide.masterContent = await this.provider.loadXML(masterFilePath)
    slide.masterIndexTable = this.indexNodes(slide.masterContent)
    slide.masterTextStyles = extractTextByPath(slide.masterContent, ["p:sldMaster", "p:txStyles"]);
    slide.masterResContent = await this.getMasterRes(masterFilePath)
    
    let layoutBgPath = this.loadLayoutBg()
    let masterBgPath = this.loadMasterBg()

    if (layoutBgPath) {
      this.layoutBg = img2Base64(await this.provider.loadArrayBuffer(layoutBgPath))
    }

    if (masterBgPath) {
      this.masterBg = img2Base64(await this.provider.loadArrayBuffer(masterBgPath))
    }

    this.slideNodes = this.slide?.content["p:sld"]["p:cSld"]["p:spTree"]
    this.slideLayoutNodes = this.slide.layoutContent["p:sldLayout"]["p:cSld"]["p:spTree"]
    slide.bgColor = this.getSlideBackgroundColor()
  }

  loadLayoutBg() {
    let resId = extractTextByPath(this.slide!.layoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr", "a:blipFill", "a:blip", "attrs", "r:embed"])
    if (!resId) {
      resId = extractTextByPath(this.slide!.layoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr", "a:blipFill", "a:blip", "attrs", "r:embed"])
    }
    
    let relationships = this.slide!.layoutResContent["Relationships"]["Relationship"]

    for (const relationship of relationships) {
      if (relationship["attrs"]["Id"] == resId) {
        return relationship["attrs"]["Target"].replace("../", "ppt/");
      }
    }

    return ""
  }
  
  loadMasterBg() {
    let resId = extractTextByPath(this.slide!.masterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr", "a:blipFill", "a:blip", "attrs", "r:embed"])
    let relationships = this.slide!.masterResContent["Relationships"]["Relationship"]

    for (const relationship of relationships) {
      if (relationship["attrs"]["Id"] == resId) {
        return relationship["attrs"]["Target"].replace("../", "ppt/");
      }
    }

    return ""
  }

  async genHtml() {
    let { slideWidth, slideHeight } = this.gprops
    let { bgColor } = this.slide!
    
    let styleName = `section-${this.index}`
    let style = new CssStyle(styleName)

    if (this.layoutBg) {
      style.addBGBase64Img(this.layoutBg)
    } else if (this.masterBg) {
      style.addBGBase64Img(this.masterBg)
    }

    style.addWidth(slideWidth!)
    style.addHeight(slideHeight!)
    style.add("background-color", "#" + bgColor!)
    style.add("background-size", "cover")
    this.gprops.addStyle(styleName, style)

    let result = `<section class="${styleName}">`
    
    let slideLayoutNodes = this.slideLayoutNodes
    for (const nodeKey in slideLayoutNodes) {
      if (nodeKey != "p:pic") {
        continue
      }
      
      if (slideLayoutNodes[nodeKey].constructor === Array) {
        for (let i = 0; i < slideLayoutNodes[nodeKey].length; i++) {
          let item = await this.processSlideNode(nodeKey, slideLayoutNodes[nodeKey][i])
          result += item
        }
      } else {
        let item = await this.processSlideNode(nodeKey, slideLayoutNodes[nodeKey])
        result += item
      }
    }

    let nodes = this.slideNodes
    for (let nodeKey in nodes) {
      if (nodes[nodeKey].constructor === Array) {
        for (let i = 0; i < nodes[nodeKey].length; i++) {
          let item = await this.processSlideNode(nodeKey, nodes[nodeKey][i])
          result += item
        }
      } else {
        let item = await this.processSlideNode(nodeKey, nodes[nodeKey])
        result += item
      }
    }

    return result + "</section>";
  }

  async getSlideRes(slidePath: string) {
    let slideResPath = slidePath.replace("slides/slide", "slides/_rels/slide") + ".rels";
    let slideResContent = await this.provider.loadXML(slideResPath)
    let relationships = slideResContent["Relationships"]["Relationship"];
    let layoutFilePath = "";
    let slideResObj: any = {}

    if (relationships.constructor === Array) {
      for (let i = 0; i < relationships.length; i++) {
        switch (relationships[i]["attrs"]["Type"]) {
          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
            layoutFilePath = relationships[i]["attrs"]["Target"].replace("../", "ppt/");
            break;
          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":
          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
          // 增加音频处理
          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio":
          default:
            slideResObj[relationships[i]["attrs"]["Id"]] = {
              "type": relationships[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
              "target": relationships[i]["attrs"]["Target"].replace("../", "ppt/")
            };
        }
      }
    } else {
      layoutFilePath = relationships["attrs"]["Target"].replace("../", "ppt/");
    }

    return {
      layoutFilePath: layoutFilePath,
      slideResContent: slideResObj,
    }
  }

  async getSlideLayoutRes(layoutFilePath: string) {
    let layoutResFilePath = layoutFilePath.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
    let layoutResContent = await this.provider.loadXML(layoutResFilePath)

    return layoutResContent
  }

  async getMasterRes(masterPath: string) {
    let masterResFilePath = masterPath.replace("slideMasters", "slideMasters/_rels") + ".rels";
    let mastertResContent = await this.provider.loadXML(masterResFilePath)

    return mastertResContent
  }

  // 从slideLayoutRes中提取出母版地址
  getSlideMasterFilePath(slideLayoutResContent: any) {
    let relationshipArray = slideLayoutResContent["Relationships"]["Relationship"];
    let masterFilename = "";
    if (relationshipArray.constructor === Array) {
      for (let i = 0; i < relationshipArray.length; i++) {
        switch (relationshipArray[i]["attrs"]["Type"]) {
          case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
            masterFilename = relationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
            break;
          default:
        }
      }
    } else {
      masterFilename = relationshipArray["attrs"]["Target"].replace("../", "ppt/");
    }

    return masterFilename
  }

  // 获取背景填充
  getSlideBackgroundColor() {
    let { content, layoutResContent, masterContent } = this.slide!
    let bgColor = this.getSolidFill(extractTextByPath(content, ["p:sld", "p:cSld", "p:bg", "p:bgPr", "a:solidFill"]));
    if (bgColor === undefined) {
      bgColor = this.getSolidFill(extractTextByPath(layoutResContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr", "a:solidFill"]));
      if (bgColor === undefined) {
        bgColor = this.getSolidFill(extractTextByPath(masterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr", "a:solidFill"]));
        if (bgColor === undefined) {
          bgColor = "FFF";
        }
      }
    }

    return bgColor;
  }

  getSolidFill(solidFill: any) {
    if (solidFill === undefined) {
      return undefined;
    }

    let color = "FFF";

    if (solidFill["a:srgbClr"] !== undefined) {
      color = extractTextByPath(solidFill["a:srgbClr"], ["attrs", "val"]);
    } else if (solidFill["a:schemeClr"] !== undefined) {
      let schemeClr = "a:" + extractTextByPath(solidFill["a:schemeClr"], ["attrs", "val"]);
      color = getSchemeColorFromTheme(this.slide!.gprops!.theme, schemeClr);
    }

    return color;
  }

  // 生成节点索引，方便后续查询node
  // 可通过id, idx, type找到节点
  indexNodes(content: any) {
    let keys = Object.keys(content);
    let spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];

    let idTable: any = {};
    let idxTable: any = {};
    let typeTable: any = {};

    for (let key in spTreeNode) {
      if (key == "p:nvGrpSpPr" || key == "p:grpSpPr") {
        continue;
      }

      let targetNode = spTreeNode[key];
      if (targetNode.constructor === Array) {
        for (let i = 0; i < targetNode.length; i++) {
          let nvSpPrNode = targetNode[i]["p:nvSpPr"];
          let id = extractTextByPath(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
          let idx = extractTextByPath(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
          let type = extractTextByPath(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

          if (id !== undefined) {
            idTable[id] = targetNode[i];
          }
          if (idx !== undefined) {
            idxTable[idx] = targetNode[i];
          }
          if (type !== undefined) {
            typeTable[type] = targetNode[i];
          }
        }
      } else {
        let nvSpPrNode = targetNode["p:nvSpPr"];
        let id = extractTextByPath(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
        let idx = extractTextByPath(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
        let type = extractTextByPath(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

        if (id !== undefined) {
          idTable[id] = targetNode;
        }
        if (idx !== undefined) {
          idxTable[idx] = targetNode;
        }
        if (type !== undefined) {
          typeTable[type] = targetNode;
        }
      }

    }

    return { "idTable": idTable, "idxTable": idxTable, "typeTable": typeTable };
  }

  async processSlideNode(nodeType: string, nodeVal: any): Promise<NodeElement | null> {
    let node: NodeElement | null = null
    switch (nodeType) {
      case "p:sp":    // Shape, Text
        node = await this.processShapeAndTextNode(nodeVal);
        break;
      case "p:cxnSp":    // Shape, Text (with connection)
        node = await this.processCxnSpNode(nodeVal);
        break;
      case "p:pic":    // Picture
        node = await this.processPicNode(nodeVal);
        break;
      case "p:graphicFrame":    // Chart, Diagram, Table
        node = await this.processGraphicFrameNode(nodeVal);
        break;
      // case "p:grpSp":    // 群組
      //   node = await this.processGroupSpNode(nodeVal);
      //   break;
      default:
        break
    }

    return node;
  }

  async processShapeAndTextNode(nodeVal: any) {
    let sp = new ShapeTextProcessor(this.provider, this.slide!, nodeVal, this.globalCssStyles, false)
    let html = await sp.genHTML()
    return html
  }

  async processCxnSpNode(nodeVal: any) {
    let sp = new ShapeTextProcessor(this.provider, this.slide!, nodeVal, this.globalCssStyles, true)
    return await sp.genHTML()
  }

  async processPicNode(nodeVal: any) {
    let picNode = new PicProcessor(this.provider, this.slide!, nodeVal, this.globalCssStyles)
    return await picNode.genHTML()
  }

  async processGraphicFrameNode(nodeVal: any) {
    let n = new GraphicProcessor(this.provider, this.slide!, nodeVal, this.globalCssStyles)
    return await n.genHTML()
  }

  async processGroupSpNode(node: any) {
    let factor = 96 / 914400;
    
    let xfrmNode = node["p:grpSpPr"]["a:xfrm"];
    let x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * factor;
    let y = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * factor;
    let chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * factor;
    let chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * factor;
    let cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * factor;
    let cy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * factor;
    let chcx = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * factor;
    let chcy = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * factor;
    let order = node["attrs"]["order"];
    let result = "<div class='block group' style='z-index: " + order + "; top: " + (y - chy) + "px; left: " + (x - chx) + "px; width: " + (cx - chcx) + "px; height: " + (cy - chcy) + "px;'>";
    
    // Procsee all child nodes
    for (let nodeKey in node) {
        if (node[nodeKey].constructor === Array) {
            for (let i=0; i<node[nodeKey].length; i++) {
                result += await this.processSlideNode(nodeKey, node[nodeKey][i]);
            }
        } else {
            result += await this.processSlideNode(nodeKey, node[nodeKey]);
        }
    }
    
    result += "</div>";
    return result;
  }
}