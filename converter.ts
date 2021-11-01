import SlideProcessor from './slideproc'
import PPTXProvider from './provider'
import { computePixel, img2Base64 } from './util'
import { GlobalProps, ThemeContent } from './model'
import fs from 'fs'
import path from 'node:path'
import { Drawer } from './drawer'

export default class PPTXConverter {
  srcFilePath: string
  outDir: string
  gprops: GlobalProps
  globalCssStyles: any
  provider: PPTXProvider
  drawer: Drawer

  constructor(srcFilePath: string, outDir: string, drawer: Drawer) {
    this.srcFilePath = srcFilePath
    this.outDir = outDir
    this.globalCssStyles = {}
    this.gprops = new GlobalProps()
    this.provider = new PPTXProvider(this.srcFilePath)
    this.drawer = drawer
  }

  async run() {
    await this.provider.init()
    let [slideWidth, slideHeight] = await this.loadSlideSize()
    let [slidePaths, slideLayouts] = await this.loadSlidesAndLayouts()

    this.gprops.slideWidth = slideWidth
    this.gprops.slideHeight = slideHeight
    this.gprops.thumbnail = await this.loadThumbImg()

    this.gprops.slidePaths = slidePaths
    this.gprops.slideLayouts = slideLayouts
    this.gprops.theme = new ThemeContent(await this.loadTheme()) 

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
    let slidesHtml = ""

    let i = 0
    for (const slide of this.gprops?.slidePaths!) {
      let processor = new SlideProcessor(
        slide, i, this.provider!,  this.gprops!, this.globalCssStyles
      )
      let content = await processor.process()
      slidesHtml += `<div class="item ${i} ${i == 0 ? "active" : ""}" >${content}</div>`
      i++

      // if (i == 1) {
      //   break
      // }
    }

    let html = this.mixContent(slidesHtml)
    fs.writeFileSync(this.getOutputName(), html)
  }

  mixContent(slidesContent: string) {
    let template = fs.readFileSync("./web/pptx.html").toString()
    let cssContent = fs.readFileSync("./web/pptx.css").toString()

    let globalCss = this.genGlobalCSS()
    let content = template.replace("{{content}}", slidesContent)
    content = content.replace("{{style}}", cssContent + " " + globalCss)
    content = content.replace("{{width}}", this.gprops!.slideWidth + "")

    return content
  }

  getOutputName(): string {
    return this.outDir + "/" +path.basename(this.srcFilePath).split(".")[0] + ".html"
}

  genGlobalCSS() {
    let cssText = "";
    for (var key in this.globalCssStyles) {
      cssText += "section ." + this.globalCssStyles[key]["name"] + "{" + this.globalCssStyles[key]["text"] + "}\n";
    }

    for (const styleName in this.gprops.globalStyles) {
      cssText += this.gprops.globalStyles[styleName].toString()
    }

    return cssText;
  }
}
