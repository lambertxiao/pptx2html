import PPTXConverter from "./converter";
import PPTXProvider from "./provider";

async function main() {
  let filepath = "/Users/lambert.xiao/workspace/mpptx2html/temp/demo.pptx"
  // let filepath = "/Users/lambert.xiao/Documents/UDI规划.pptx"
  let provider = new PPTXProvider(filepath)
  await provider.provide()

  let converter = new PPTXConverter(provider)
  await converter.loadPPTX()
}

main()
