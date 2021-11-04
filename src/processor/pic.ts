import { PicNode } from '../model';
import { extractText, getImgMimeType, toBase64ImgLink } from '../util';
import NodeProcessor from './processor';

export default class PicProcessor extends NodeProcessor {

  async process() {
    let node = this.node
    let name = extractText(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"])
    let order = node["attrs"]["order"];
    let rid = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
    let imgPath = ""

    if (node["__location"] == "layout") {
      imgPath = this.slide.getTargetFromLayout(rid)
    } else if (node["__location"] == "slide") {
      imgPath = this.slide.getTargetFromSlide(rid)
    }

    let xfrmNode = node["p:spPr"]["a:xfrm"];
    let prst = node["p:spPr"]["a:prstGeom"]["attrs"]["prst"]
    let borderRadius: number = 0
    // 圆角矩形xml里没有给出具体的边弧度
    if (prst == "roundRect") {
      borderRadius = 40
    }

    let { top, left } = this.getPosition(xfrmNode, undefined, undefined)
    let { width, height } = this.getSize(xfrmNode, undefined, undefined)

    let mimeType = getImgMimeType(imgPath)
    let imgArrayBuffer = await this.provider.loadArrayBuffer(imgPath)
    let img = toBase64ImgLink(mimeType, imgArrayBuffer!)

    let pn: PicNode = {
      name: name,
      eleType: "pic",
      width: width,
      height: height,
      zindex: order,
      top: top,
      left: left,
      imgUrl: img,
      mimeType: mimeType,
      borderRadius: borderRadius,
    }

    return pn
  }
}
