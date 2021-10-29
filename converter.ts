import SlideProcessor from './slide_processor'
import PPTXProvider from './provider'
import { computePixel, img2Base64 } from './util'
import { GlobalProps } from './props'
import fs from 'fs'

export default class PPTXConverter {
  gprops?: GlobalProps
  globalCssStyles: any

  constructor(private readonly provider: PPTXProvider) {
    this.globalCssStyles = {}
  }

  async loadPPTX() {
    let gprops = new GlobalProps()
    let [slideWidth, slideHeight] = await this.loadSlideSize()
    gprops.slideWidth = slideWidth
    gprops.slideHeight = slideHeight
    gprops.thumbnail = await this.loadThumbImg()

    let [slidePaths, slideLayouts] = await this.loadSlidesAndLayouts()
    gprops.slidePaths = slidePaths
    gprops.slideLayouts = slideLayouts

    gprops.theme = await this.loadTheme()
    this.gprops = gprops

    await this.processSlides()
  }

  async loadThumbImg() {
    return img2Base64(await this.provider.loadArrayBuffer("docProps/thumbnail.jpeg"))
  }

  // 读取[Content_Types].xml，解析出slides和slideLayouts
  async loadSlidesAndLayouts() {
    let contentTypes = await this.provider.loadXML("[Content_Types].xml")
    let subObj = contentTypes["Types"]["Override"]
    let slidesLocArray = []
    let slideLayoutsLocArray = []

    for (let i = 0; i < subObj.length; i++) {
      switch (subObj[i]["attrs"]["ContentType"]) {
        case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
          slidesLocArray.push(subObj[i]["attrs"]["PartName"].substr(1))
          break
        case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
          slideLayoutsLocArray.push(subObj[i]["attrs"]["PartName"].substr(1))
          break
        default:
          break
      }
    }

    return [slidesLocArray, slideLayoutsLocArray]

  }

  // 获取幻灯片宽高
  async loadSlideSize() {
    let content = await this.provider.loadXML("ppt/presentation.xml")
    let sldSzAttrs = content["p:presentation"]["p:sldSz"]["attrs"]
    let slideWidth = computePixel(sldSzAttrs["cx"])
    let slideHeight = computePixel(sldSzAttrs["cy"])

    return [slideWidth, slideHeight]
  }

  async loadTheme() {
    let prenContent = await this.provider.loadXML("ppt/_rels/presentation.xml.rels")
    let relationships = prenContent["Relationships"]["Relationship"]
    let themeURI = undefined;

    if (relationships.constructor === Array) {
      for (let i = 0; i < relationships.length; i++) {
        if (relationships[i]["attrs"]["Type"] === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
          themeURI = relationships[i]["attrs"]["Target"];
          break;
        }
      }
    } else if (relationships["attrs"]["Type"] === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
      themeURI = relationships["attrs"]["Target"];
    }

    if (themeURI === undefined) {
      throw Error("Can't open theme file.");
    }

    return await this.provider.loadXML("ppt/" + themeURI)
  }

  async processSlides() {
    let html = ""
    let i = 0
    for (const slide of this.gprops?.slidePaths!) {
      let processor = new SlideProcessor(this.provider!, slide, this.gprops!, this.globalCssStyles)
      let content = await processor.process()
      html += `<div class="item ${i} ${i == 0 ? "active" :""}" >${content}</div>`
      i++
    }

    let template = fs.readFileSync("./web/pptx.html").toString()
    let cssContent = fs.readFileSync("./web/pptx.css").toString()
    let globalCss = this.genGlobalCSS()
    let content = template.replace("{{content}}", html)
    content = content.replace("{{style}}", cssContent + " " + globalCss)
    content = content.replace("{{width}}", this.gprops!.slideWidth + "")

    fs.writeFileSync("./a.html", content)
  }

  genGlobalCSS() {
    var cssText = "";
    for (var key in this.globalCssStyles) {
      cssText += "section ." + this.globalCssStyles[key]["name"] + "{" + this.globalCssStyles[key]["text"] + "}\n";
    }
    return cssText;
  }
}
