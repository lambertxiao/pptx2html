import PPTXConverter from "./converter";
import PPTXProvider from "./provider";
import { program } from "commander"
import path from "node:path";

async function main() {
  program
    .option('-s, --src <string>', 'src pptx file')
    .option('-o, --outdir <string>', 'output dir')

  program.parse(process.argv)
  const options = program.opts()
  let srcFile = path.resolve(options.src)
  let outDir = path.resolve(options.outdir)

  let converter = new PPTXConverter(srcFile, outDir)
  await converter.run()
}

main()
