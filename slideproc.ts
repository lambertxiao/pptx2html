import PPTXProvider from './provider';
import { GlobalProps, NodeElement, NodeElementGroup, SingleSlide, SlideView } from './model';
import { extractTextByPath, getSchemeColorFromTheme, img2Base64, printObj } from './util';
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
  ) {
    this.gprops = gprops
  }

  async process() {
    await this.prepare()
    return await this.genSlideView()
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

  async genSlideView() {
    let sv = new SlideView()

    let { slideWidth, slideHeight } = this.gprops
    sv.width = slideWidth
    sv.height = slideHeight

    let { bgColor } = this.slide!
    sv.bgColor = bgColor

    if (this.layoutBg) {
      sv.bgImgData = this.layoutBg
    } else if (this.masterBg) {
      sv.bgImgData = this.masterBg
    }

    sv.bgColor = bgColor

    let slideLayoutNodes = this.slideLayoutNodes
    for (const nodeKey in slideLayoutNodes) {
      if (nodeKey != "p:pic") {
        continue
      }

      if (slideLayoutNodes[nodeKey].constructor === Array) {
        for (let i = 0; i < slideLayoutNodes[nodeKey].length; i++) {
          let node = slideLayoutNodes[nodeKey][i]
          node["__location"] = "layout"
          let item = await this.processSlideNode(nodeKey, node)
          if (item) {
            sv.addLayoutNode(item)
          }
        }
      } else {
        let node = slideLayoutNodes[nodeKey]
        node["__location"] = "layout"
        let item = await this.processSlideNode(nodeKey, node)
        if (item) {
          sv.addLayoutNode(item)
        }
      }
    }

    let nodes = this.slideNodes
    for (let nodeKey in nodes) {
      if (nodes[nodeKey].constructor === Array) {
        for (let i = 0; i < nodes[nodeKey].length; i++) {
          let node = nodes[nodeKey][i]
          node["__location"] = "slide"
          let item = await this.processSlideNode(nodeKey, node)
          if (item) {
            sv.addSlideNode(item)
          }
        }
      } else {
        let node = nodes[nodeKey]
        node["__location"] = "slide"
        let item = await this.processSlideNode(nodeKey, node)
        if (item) {
          sv.addSlideNode(item)
        }
      }
    }

    return sv;
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
          bgColor = "#FFF";
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

  async processSlideNode(nodeType: string, nodeVal: any): Promise<NodeElement | NodeElementGroup | null> {
    let node: NodeElement | NodeElementGroup | null = null
    switch (nodeType) {
      // Shape, Text
      case "p:sp":
        node = await this.processShapeAndTextNode(nodeVal);
        break
      // Shape, Text (with connection)
      case "p:cxnSp":
        node = await this.processCxnSpNode(nodeVal);
        break
      // Picture
      case "p:pic":
        node = await this.processPicNode(nodeVal);
        break
      case "p:graphicFrame":
        // Chart, Diagram, Table
        node = await this.processGraphicFrameNode(nodeVal);
        break;
      case "p:grpSp":    // 群組
        console.log("parse 群组")
        node = await this.processGroupSpNode(nodeVal);
        break;
      default:
        break
    }

    return node;
  }

  async processShapeAndTextNode(nodeVal: any) {
    let sp = new ShapeTextProcessor(this.provider, this.slide!, nodeVal, false)
    let html = await sp.process()
    return html
  }

  async processCxnSpNode(nodeVal: any) {
    let sp = new ShapeTextProcessor(this.provider, this.slide!, nodeVal, true)
    return await sp.process()
  }

  async processPicNode(nodeVal: any) {
    let picNode = new PicProcessor(this.provider, this.slide!, nodeVal)
    return await picNode.process()
  }

  async processGraphicFrameNode(nodeVal: any) {
    let n = new GraphicProcessor(this.provider, this.slide!, nodeVal)
    return await n.process()
  }

  async processGroupSpNode(node: any) {
    let group = new NodeElementGroup()
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

    group.zindex = order
    group.top = y - chy
    group.left = x - chx
    group.width = cx - chcx
    group.height = cy - chcy

    // Procsee all child nodes
    for (let nodeKey in node) {
        if (node[nodeKey].constructor === Array) {
            for (let i=0; i<node[nodeKey].length; i++) {
              let n = await this.processSlideNode(nodeKey, node[nodeKey][i])
              if (n) {
                group.nodes.push(n)
              }
            }
        } else {
          let n = await this.processSlideNode(nodeKey, node[nodeKey])
          if (n) {
            group.nodes.push(n)
          }
        }
    }

    return group
  }
}
