import { extractFileExtension, img2Base64 } from '../util';
import NodeResolver from './base';

export default class PicNode extends NodeResolver {

  async genHTML() {
    let node = this.node
    let order = node["attrs"]["order"];
    let rid = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
    let imgName = this.slide.resContent[rid]["target"];
    let imgFileExt = extractFileExtension(imgName).toLowerCase();
    let imgArrayBuffer = await this.provider.loadArrayBuffer(imgName)
    let mimeType = "";
    let xfrmNode = node["p:spPr"]["a:xfrm"];

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
    return "<div class='block content' style='" + this.getPosition(xfrmNode, undefined, undefined) + this.getSize(xfrmNode, undefined, undefined) +
      " z-index: " + order + ";" +
      "'><img src=\"data:" + mimeType + ";base64," + img2Base64(imgArrayBuffer) + "\" style='width: 100%; height: 100%'/></div>";
  }
}