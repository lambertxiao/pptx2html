import PPTXConverter from "../pptx_converter";
import PPTXProvider from "../pptx_provider";

async function main() {
  let filepath = "/Users/lambert.xiao/workspace/pptxtohtml/example/demo.pptx"

  let provider = new PPTXProvider(filepath)
  await provider.provide()

  let converter = new PPTXConverter(provider)
  await converter.loadPPTX()
}

main()
