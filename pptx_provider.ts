import JSZip from 'jszip';
import fs from 'fs';
import XMLTransform from './xml_transform'

export default class PPTXProvider {
  pptxData?: JSZip

  constructor(private readonly pptxPath: string) {}

  async provide() {
    let data = fs.readFileSync(this.pptxPath)
    let zip = new JSZip()
    this.pptxData = await zip.loadAsync(data)
  }

  async loadXML(path: string) {
    let strContent = await this.pptxData?.file(path)?.async("string")
    let xt = new XMLTransform()
    return xt.toXML(strContent!)
  }

  async loadBlob(path: string) {
    return await this.pptxData?.file(path)?.async("blob")
  }

  async loadArrayBuffer(path: string) {
    return await this.pptxData?.file(path)?.async("arraybuffer")
  }
}
