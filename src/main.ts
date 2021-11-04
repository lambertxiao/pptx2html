import PPTXConverter from "./converter";
import { program } from "commander"
import path from "node:path";
import { HtmlDrawer } from "./drawer_html";
import fs from "fs"

async function main() {
  program
    .option('-s, --src <string>', 'src pptx file')
    .option('-o, --outdir <string>', 'output dir')
    .option('-p, --page <number>', 'specified page')

  program.parse(process.argv)
  const options = program.opts()
  let srcFile = path.resolve(options.src)
  let outDir = path.resolve(options.outdir)

  let converter = new PPTXConverter(srcFile)
  let slideViews = await converter.convert(options.page)
  let templateHtml = path.resolve("../web/pptx.html")
  let templateCss = path.resolve("../web/pptx.css")
  let drawer = new HtmlDrawer(templateHtml, templateCss)
  let html = drawer.draw(slideViews)

  let outFile = outDir + "/" +path.basename(srcFile).split(".")[0] + ".html"
  fs.writeFileSync(outFile, html)
}

main()
