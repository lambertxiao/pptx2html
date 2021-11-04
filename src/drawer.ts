import { NodeElement, PicNode, ShapeNode, SlideView, SpanNode, TableNode, TextNode } from "./model";
import fs from 'fs'
import { printObj } from "./util";

export interface Drawer {
  draw(nodes: SlideView[]): string
}

export class CssStyle {

  content: { [key: string]: string } = {}

  constructor(private readonly name: string) {
    this.name = name
  }

  add(key: string, val: string) {
    this.content[key] = val
  }

  addWidth(val: number) {
    this.add("width", val + "px")
  }

  addHeight(val: number) {
    this.add("height", val + "px")
  }

  addBGBase64Img(val: string) {
    this.add("background-image", `url(data:image/png;base64,${val})`)
  }

  toString() {
    let s = `.${this.name} {`
    for (const key in this.content) {
      s += `${key}: ${this.content[key]};`
    }

    s += "}"
    return s
  }
}

export class HtmlDrawer implements Drawer {

  cssStyles: CssStyle[] = []

  constructor(
    private readonly templateHtml: string,
    private readonly templateCss: string,
  ) { }

  draw(slideViews: SlideView[]): string {
    let html = ""

    for (let i = 0; i < slideViews.length; i++) {
      let sv = slideViews[i]
      let styleName = "section-" + i
      let style = new CssStyle(styleName)
      style.addWidth(sv.width!)
      style.addHeight(sv.height!)
      style.add("background-size", "cover")

      if (sv.bgImgData) {
        style.addBGBase64Img(sv.bgImgData)
      } else {
        style.add("background-color", sv.bgColor!)
      }

      this.cssStyles.push(style)
      let section = `<section class="${styleName}">`
      let nodes: NodeElement[] = [...sv.layoutNodes, ...sv.slideNodes]

      for (const ln of nodes) {
        switch (ln.eleType) {
          case "text":
            let textNode = <TextNode>ln
            section += this.drawTextNode(textNode)
            break
          case "shape":
            let shapeNode = <ShapeNode>ln
            section += this.drawShapeNode(shapeNode)
            break
          case "table":
            let tableNode = <TableNode>ln
            section += this.drawTableNode(tableNode)
            break
          case "pic":
            let picNode = <PicNode>ln
            section += this.drawPicNode(picNode)
            break
          // case "chart":
          //   break
          // case "diagram":
          //   break
        }
      }

      section += `</section>`
      html += section
    }

    html = this.mixContent(html)
    return html
  }

  getNodeBasicStyle(node: NodeElement) {
    let styles: string[] = []
    styles.push(`width: ${node.width}px;`)

    if (node.height) {
      styles.push(`height: ${node.height}px;`)
    }

    if (node.top) {
      styles.push(`top: ${node.top}px;`)
    }

    if (node.left) {
      styles.push(`left: ${node.left}px;`)
    }

    if (node.zindex) {
      styles.push(`z-index: ${node.zindex};`)
    }

    return styles.join('')
  }

  drawPicNode(node: PicNode) {
    let styles = this.getNodeBasicStyle(node)
    let radius = ""
    if (node.borderRadius) {
      radius = `border-radius: ${node.borderRadius}px`
    }

    let html = `
      <div class="block content" style="${styles}">
        <img src="${node.imgUrl}" style="width: 100%; height: 100%; ${radius}"/>
      </div>
    `

    return html
  }

  drawTableNode(node: TableNode) {
    let styles = this.getNodeBasicStyle(node)
    let html = `<table style="${styles}; border: none;">`

    for (const row of node.rows) {
      html += `<tr>`
      for (const col of row.cols) {
        let text = this.drawColTextNode(col.text!)
        if (!text) {
          continue
        }

        if (col.rowSpan) {
          html += `<td rowspan="${col.rowSpan}">${text}</td>`
        } else if (col.colSpan) {
          html += `<td colspan="${col.colSpan}">${text}</td>`
        } else {
          html += `<td>${text}</td>`
        }
      }

      html += `</tr>`
    }

    html += `</table>`
    return html
  }

  drawShapeNode(node: ShapeNode): string {
    let styles = this.getNodeBasicStyle(node)
    let html = `<svg class="drawing" style="${styles}">`
    let fillColor = node.bgColor ? node.bgColor : "none"
    let borderColor = node.border?.color ? node.border?.color : "none"

    console.log(node.bgImg)

    switch (node.shapeType) {
      case "accentBorderCallout1":
      case "accentBorderCallout2":
      case "accentBorderCallout3":
      case "accentCallout1":
      case "accentCallout2":
      case "accentCallout3":
      case "actionButtonBackPrevious":
      case "actionButtonBeginning":
      case "actionButtonBlank":
      case "actionButtonDocument":
      case "actionButtonEnd":
      case "actionButtonForwardNext":
      case "actionButtonHelp":
      case "actionButtonHome":
      case "actionButtonInformation":
      case "actionButtonMovie":
      case "actionButtonReturn":
      case "actionButtonSound":
      case "arc":
      case "bevel":
      case "blockArc":
      case "borderCallout1":
      case "borderCallout2":
      case "borderCallout3":
      case "bracePair":
      case "bracketPair":
      case "callout1":
      case "callout2":
      case "callout3":
      case "can":
      case "chartPlus":
      case "chartStar":
      case "chartX":
      case "chevron":
      case "chord":
      case "cloud":
      case "cloudCallout":
      case "corner":
      case "cornerTabs":
      case "cube":
      case "decagon":
      case "diagStripe":
      case "diamond":
      case "dodecagon":
      case "donut":
      case "doubleWave":
      case "downArrowCallout":
      case "ellipseRibbon":
      case "ellipseRibbon2":
      case "flowChartAlternateProcess":
      case "flowChartCollate":
      case "flowChartConnector":
      case "flowChartDecision":
      case "flowChartDelay":
      case "flowChartDisplay":
      case "flowChartDocument":
      case "flowChartExtract":
      case "flowChartInputOutput":
      case "flowChartInternalStorage":
      case "flowChartMagneticDisk":
      case "flowChartMagneticDrum":
      case "flowChartMagneticTape":
      case "flowChartManualInput":
      case "flowChartManualOperation":
      case "flowChartMerge":
      case "flowChartMultidocument":
      case "flowChartOfflineStorage":
      case "flowChartOffpageConnector":
      case "flowChartOnlineStorage":
      case "flowChartOr":
      case "flowChartPredefinedProcess":
      case "flowChartPreparation":
      case "flowChartProcess":
      case "flowChartPunchedCard":
      case "flowChartPunchedTape":
      case "flowChartSort":
      case "flowChartSummingJunction":
      case "flowChartTerminator":
      case "folderCorner":
      case "frame":
      case "funnel":
      case "gear6":
      case "gear9":
      case "halfFrame":
      case "heart":
      case "heptagon":
      case "hexagon":
      case "homePlate":
      case "horizontalScroll":
      case "irregularSeal1":
      case "irregularSeal2":
      case "leftArrow":
      case "leftArrowCallout":
      case "leftBrace":
      case "leftBracket":
      case "leftRightArrowCallout":
      case "leftRightRibbon":
      case "irregularSeal1":
      case "lightningBolt":
      case "lineInv":
      case "mathDivide":
      case "mathEqual":
      case "mathMinus":
      case "mathMultiply":
      case "mathNotEqual":
      case "mathPlus":
      case "moon":
      case "nonIsoscelesTrapezoid":
      case "noSmoking":
      case "octagon":
      case "parallelogram":
      case "pentagon":
      case "pie":
      case "pieWedge":
      case "plaque":
      case "plaqueTabs":
      case "plus":
      case "quadArrowCallout":
      case "ribbon":
      case "ribbon2":
      case "rightArrowCallout":
      case "rightBrace":
      case "rightBracket":
      case "round1Rect":
      case "round2DiagRect":
      case "round2SameRect":
      case "rtTriangle":
      case "smileyFace":
      case "snip1Rect":
      case "snip2DiagRect":
      case "snip2SameRect":
      case "snipRoundRect":
      case "squareTabs":
      case "star10":
      case "star12":
      case "star16":
      case "star24":
      case "star32":
      case "star4":
      case "star5":
      case "star6":
      case "star7":
      case "star8":
      case "sun":
      case "teardrop":
      case "trapezoid":
      case "upArrowCallout":
      case "upDownArrowCallout":
      case "verticalScroll":
      case "wave":
      case "wedgeEllipseCallout":
      case "wedgeRectCallout":
      case "wedgeRoundRectCallout":
      case "rect":
        html +=
        `<rect
          x=0 y=0
          width=${node.ShapeWidth} height=${node.ShapeHeight}
          fill="${fillColor}"
          stroke="${borderColor}"
          stroke-width="${node.border!.width}"
          stroke-dasharray="${node.border!.strokeDasharray}"
          />`
        break;
      case "ellipse":
        html +=
        `<ellipse
          cx="${node.ShapeWidth! / 2} cy="${node.ShapeHeight! / 2}"
          rx="${node.ShapeWidth! / 2} cy="${node.ShapeHeight! / 2}"
          fill="${fillColor}"
          stroke="${borderColor}"
          stroke-width="${node.border!.width}"
          stroke-dasharray="${node.border!.strokeDasharray}"
        `
        break;
      case "roundRect":
        html +=
        `<rect
          x=0 y=0
          width=${node.ShapeWidth} height=${node.ShapeHeight}
          rx="7" ry="7"
          fill="${fillColor}"
          stroke="${borderColor}"
          stroke-width="${node.border!.width}"
          stroke-dasharray="${node.border!.strokeDasharray}"
          />`
        break;
      case "bentConnector2":    // 直角 (path)
        let d = "";
        if (node.isFlipV) {
          d = "M 0 " + node.ShapeWidth + " L " + node.ShapeHeight + " " + node.ShapeWidth + " L " + node.ShapeHeight + " 0";
        } else {
          d = "M " + node.ShapeWidth + " 0 L " + node.ShapeWidth + " " + node.ShapeHeight + " L 0 " + node.ShapeHeight;
        }

        html +=
        `
        <path
          d="${d}"
          stroke="${borderColor}"
          stroke-width="${node.border!.width}"
          stroke-dasharray="${node.border!.strokeDasharray}"
          fill="${fillColor}"
        />
        `
        break;
      case "line":
      case "straightConnector1":
      case "bentConnector3":
      case "bentConnector4":
      case "bentConnector5":
      case "curvedConnector2":
      case "curvedConnector3":
      case "curvedConnector4":
      case "curvedConnector5":
        if (node.isFlipV) {
          html +=
          `<line
            x1="${node.ShapeWidth}" y1='0'
            x2='0' y2="${node.ShapeHeight}"
            stroke="${borderColor}"
            stroke-width="${node.border!.width}"
            stroke-dasharray="${node.border!.strokeDasharray}"
          />
          `;
        } else {
          html +=
          `<line
            x1="0" y1="0"
            x2="${node.ShapeWidth}" y2="${node.ShapeHeight}"
            stroke="${borderColor}"
            stroke-width="${node.border!.width}"
            stroke-dasharray="${node.border!.strokeDasharray}"
          />
          `;
        }
        break;
      case "rightArrow":
        html += `
        <defs>
          <marker
            id="markerTriangle"
            viewBox="0 0 10 10"
            refX="1" refY="5"
            markerWidth="2.5"
            markerHeight="2.5"
            orient="auto-start-reverse"
            markerUnits="strokeWidth">
            <path d="M 0 0 L 10 5 L 0 10 z"/>
          </marker>
        </defs>
        <line
          x1="0" y1="${node.ShapeHeight! / 2}"
          x2="${node.ShapeWidth! - 15}" y2="${node.ShapeHeight! / 2}"
          stroke="${borderColor}"
          stroke-width="${node.border!.width}"
          stroke-dasharray="${node.border!.strokeDasharray}"
          marker-end='url(#markerTriangle)' />`
        break;
      case "downArrow":
        html += `
        <defs>
          <marker
            id="markerTriangle"
            viewBox="0 0 10 10"
            refX="1" refY="5"
            markerWidth="2.5"
            markerHeight="2.5"
            orient="auto-start-reverse"
            markerUnits="strokeWidth">
            <path d="M 0 0 L 10 5 L 0 10 z"/>
          </marker>
        </defs>
        <line
          x1="${node.ShapeWidth! / 2}" y1="${node.ShapeHeight! / 2}"
          x2="${node.ShapeWidth! / 2}" y2="${node.ShapeHeight! - 15}"
          stroke="${borderColor}"
          stroke-width="${node.border!.width}"
          stroke-dasharray="${node.border!.strokeDasharray}"
          marker-end='url(#markerTriangle)' />`
        break;
      case "bentArrow":
      case "bentUpArrow":
      case "stripedRightArrow":
      case "quadArrow":
      case "circularArrow":
      case "swooshArrow":
      case "leftRightArrow":
      case "leftRightUpArrow":
      case "leftUpArrow":
      case "leftCircularArrow":
      case "notchedRightArrow":
      case "curvedDownArrow":
      case "curvedLeftArrow":
      case "curvedRightArrow":
      case "curvedUpArrow":
      case "upDownArrow":
      case "upArrow":
      case "uturnArrow":
      case "leftRightCircularArrow":
        break;
      case "triangle":
        break;
      case undefined:
        break
      default:
        console.warn("Undefine shape type.");
    }

    if (node.textNode) {
      let content = this.drawTextNode(node.textNode!)
      if (node.bgColor) {
        styles += `background-color: ${node.bgColor};`
      }

      if (node.border) {
        let border = node.border
        styles += `border: ${border.width}${border.widthUnit} ${border.type} ${border.color};`
      }

      if (!node.shapeType) {
        html += `
        <div class="block content" style="${styles}">
          ${content}
        </div>
        `
      } else {
        html += content
      }
    }

    html += `</svg>`

    return html
  }

  drawColTextNode(node: TextNode): string {
    let html = ""

    if (node.content) {
      let content = this.drawSpanNode(node.content)
      html += content
    }

    for (const ph of node.paragraphNodes) {
      let spans = ""
      for (const span of ph.spans) {
        spans += this.drawSpanNode(span)
      }

      if (spans) {
        html += `<div class="${node.styleClass}">${spans}</div>`
      }
    }

    if (!html) {
      return ""
    }

    return `<div>${html}</div>`
  }

  drawTextNode(node: TextNode): string {
    let styles: string[] = []

    if (node.color) {
      styles.push(`color: ${node.color};`)
    }

    if (node.fontFamily) {
      styles.push(`font-family: ${node.fontFamily};`)
    }

    if (node.fontSize) {
      styles.push(`font-size: ${node.fontSize};`)
    }

    if (node.width) {
      styles.push(`width: ${node.width}px;`)
    }
    if (node.height) {
      styles.push(`height: ${node.height}px;`)
    }

    if (node.top) {
      styles.push(`top: ${node.top}px;`)
    }

    if (node.left) {
      styles.push(`left: ${node.left}px;`)
    }

    if (node.zindex) {
      styles.push(`z-index: ${node.zindex};`)
    }

    let strStyles = styles.join("")

    let html = `<div class="block" style="${strStyles}">`

    if (node.content) {
      let content = this.drawSpanNode(node.content)
      html += content
    }

    for (const ph of node.paragraphNodes) {
      html += `<div class="${node.styleClass}">`
      for (const span of ph.spans) {
        html += this.drawSpanNode(span)
      }
      html += "</div>"
    }

    html += `</div>`
    return html
  }

  drawSpanNode(node: SpanNode) {
    if (!node.content) {
      return ""
    }

    let span = node
    let styles: string[] = []
    if (span.color) {
      styles.push(`color: ${span.color};`)
    }

    if (span.fontFamily) {
      styles.push(`font-family: ${span.fontFamily};`)
    }

    if (span.fontSize) {
      if (typeof (span.fontSize) == 'string') {
        styles.push(`font-size: ${span.fontSize};`)
      } else {
        styles.push(`font-size: ${span.fontSize}${span.fontSizeUnit};`)
      }
    }

    if (span.fontStyle) {
      styles.push(`font-style: ${span.fontStyle};`)
    }

    if (span.textDecoration) {
      styles.push(`text-decoration: ${span.textDecoration};`)
    }

    let strStyles = styles.join("")
    let c = `<span class="text-block" style="${strStyles}">${node.content}</span>`

    return c
  }

  genGlobalCSS() {
    let css = ""
    for (const style of this.cssStyles) {
      css += style.toString() + "\n"
    }

    return css
  }

  mixContent(slidesContent: string) {
    let template = fs.readFileSync(this.templateHtml).toString()
    let cssContent = fs.readFileSync(this.templateCss).toString()

    let globalCss = this.genGlobalCSS()
    let content = template.replace("{{content}}", slidesContent)
    content = content.replace("{{style}}", cssContent + globalCss)
    // content = content.replace("{{width}}", this.gprops!.slideWidth + "")

    return content
  }
}
