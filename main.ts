import PPTXConverter from "./pptx_converter";
import PPTXProvider from "./pptx_provider";

async function main() {
  let filepath = "/Users/star/workspace/ppt2html/temp/kj.pptx"

  let provider = new PPTXProvider(filepath)
  await provider.provide()

  let converter = new PPTXConverter(provider)
  await converter.loadPPTX()
}

main()
