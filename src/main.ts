import PPTXConverter from "./converter";
import { program } from "commander"
import path from "node:path";
import { HtmlDrawer } from "./drawer_html";
import fs from "fs"

async function main() {
  program
    .option('-s, --src <string>', 'src pptx file')
    .option('-o, --outdir <string>', 'output dir')
    .option('-fn --outFilename <string>', 'out file name')
    .option('-p, --page <number>', 'specified page')

  program.parse(process.argv)
  const options = program.opts()
  let srcFile = path.resolve(options.src)
  let outDir = path.resolve(options.outdir)
  let outFileName = options.outFilename

  let converter = new PPTXConverter(srcFile)
  let slideViews = await converter.convert(options.page)
  let templateHtml = __dirname + "/../web/pptx.html"
  let templateCss = __dirname + "/../web/pptx.css"
  let drawer = new HtmlDrawer(templateHtml, templateCss)
  let html = drawer.draw(slideViews)

  let name = ""
  if (outFileName) {
    name = outFileName
  } else {
    name = path.basename(srcFile).split(".")[0] + ".html"
  }

  let outFile = outDir + "/" + name
  fs.writeFileSync(outFile, html)
}

main()
