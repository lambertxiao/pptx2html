import { PicNode } from '../model';
import { extractFileExtension, img2Base64 } from '../util';
import NodeProcessor from './processor';

export default class PicProcessor extends NodeProcessor {

  async genHTML() {
    let node = this.node
    let order = node["attrs"]["order"];
    let rid = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
    let imgName = this.slide.getTargetResByID(rid)
    let imgFileExt = extractFileExtension(imgName).toLowerCase();
    let imgArrayBuffer = await this.provider.loadArrayBuffer(imgName)
    let mimeType = "";
    let xfrmNode = node["p:spPr"]["a:xfrm"];
    let prst = node["p:spPr"]["a:prstGeom"]["attrs"]["prst"]

    switch (imgFileExt) {
      case "jpg":
      case "jpeg":
        mimeType = "image/jpeg";
        break;
      case "png":
        mimeType = "image/png";
        break;
      case "gif":
        mimeType = "image/gif";
        break;
      case "emf": // Not native support
        mimeType = "image/x-emf";
        break;
      case "wmf": // Not native support
        mimeType = "image/x-wmf";
        break;
      default:
        mimeType = "image/*";
    }

    let borderRadius: number = 0
    // 圆角矩形xml里没有给出具体的边弧度
    if (prst == "roundRect") {
      borderRadius = 40
    }

    let { top, left } = this.getPosition(xfrmNode, undefined, undefined)
    let { width, height } = this.getSize(xfrmNode, undefined, undefined)
    let img = `data:${mimeType};base64,${img2Base64(imgArrayBuffer)}`

    let pn: PicNode = {
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
