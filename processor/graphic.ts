import { table } from 'console';
import { randomInt } from 'crypto';
import { ChartNode, DiagramNode, TableCol, TableNode, TableRow, TextNode } from '../model';
import { extractTextByPath } from '../util';
import NodeProcessor from './processor';

export default class GraphicProcessor extends NodeProcessor {

  async genHTML() {
    let node = this.node
    let graphicTypeUri = extractTextByPath(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);

    switch (graphicTypeUri) {
      case "http://schemas.openxmlformats.org/drawingml/2006/table":
        return this.genTable(node);
      case "http://schemas.openxmlformats.org/drawingml/2006/chart":
        return new ChartNode()
      case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
        return new DiagramNode()
      default:
        return null
    }
  }

  genTable(node: any) {
    let order = node["attrs"]["order"];
    let _tableNode = extractTextByPath(node, ["a:graphic", "a:graphicData", "a:tbl"]);
    let xfrmNode = extractTextByPath(node, ["p:xfrm"]);
    let { width, height } = this.getSize(xfrmNode, undefined, undefined)
    let { top, left } = this.getPosition(xfrmNode, undefined, undefined)
    let tableNode = new TableNode()
    tableNode.top = top
    tableNode.left = left
    tableNode.width = width
    tableNode.height = height
    tableNode.zindex = order

    let trNodes = _tableNode["a:tr"];

    if (trNodes.constructor === Array) {
      for (let i = 0; i < trNodes.length; i++) {
        let row = new TableRow()
        let tcNodes = trNodes[i]["a:tc"];

        if (tcNodes.constructor === Array) {
          for (let j = 0; j < tcNodes.length; j++) {
            console.log(tcNodes[j]["a:txBody"])
            let text = this.genTextBody(tcNodes[j]["a:txBody"], "");
            let rowSpan = extractTextByPath(tcNodes[j], ["attrs", "rowSpan"]);
            let colSpan = extractTextByPath(tcNodes[j], ["attrs", "gridSpan"]);
            let vMerge = extractTextByPath(tcNodes[j], ["attrs", "vMerge"]);
            let hMerge = extractTextByPath(tcNodes[j], ["attrs", "hMerge"]);

            let col = new TableCol()
            col.rowSpan = rowSpan
            col.colSpan = colSpan
            col.text = text

            row.cols.push(col)
          }
        } else {
          let col = new TableCol()
          col.text = this.genTextBody(tcNodes["a:txBody"], "")
          row.cols.push(col)
        }

        tableNode.rows.push(row)
      }
    } else {
      let row = new TableRow()
      let tcNodes = trNodes["a:tc"]
      let col = new TableCol()

      if (tcNodes.constructor === Array) {
        for (let j = 0; j < tcNodes.length; j++) {
          let tn = this.genTextBody(tcNodes[j]["a:txBody"], "")
          col.text = tn
          row.cols.push(col)
        }
      } else {
        let tn = this.genTextBody(tcNodes["a:txBody"], "");
        col.text = tn
        row.cols.push(col)
      }

      tableNode.rows.push(row);
    }

    console.log(tableNode)

    return tableNode;
  }

  genChart(node: any) {
    let chartID = randomInt(2<<10)
    let order = node["attrs"]["order"];
    let xfrmNode = extractTextByPath(node, ["p:xfrm"]);
    let result = "<div id='chart" + chartID + "' class='block content' style='" +
      this.getPosition(xfrmNode, undefined, undefined) + this.getSize(xfrmNode, undefined, undefined) +
      " z-index: " + order + ";'></div>";

    let rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
    let refName = this.slide.resContent[rid]["target"];
    let content = this.provider.loadXML(refName)
    let plotArea = extractTextByPath(content, ["c:chartSpace", "c:chart", "c:plotArea"]);

    let chartData = null;
    for (let key in plotArea) {
      switch (key) {
        case "c:lineChart":
          chartData = {
            "type": "createChart",
            "data": {
              "chartID": "chart" + chartID,
              "chartType": "lineChart",
              "chartData": this.extractChartData(plotArea[key]["c:ser"])
            }
          };
          break;
        case "c:barChart":
          chartData = {
            "type": "createChart",
            "data": {
              "chartID": "chart" + chartID,
              "chartType": "barChart",
              "chartData": this.extractChartData(plotArea[key]["c:ser"])
            }
          };
          break;
        case "c:pieChart":
          chartData = {
            "type": "createChart",
            "data": {
              "chartID": "chart" + chartID,
              "chartType": "pieChart",
              "chartData": this.extractChartData(plotArea[key]["c:ser"])
            }
          };
          break;
        case "c:pie3DChart":
          chartData = {
            "type": "createChart",
            "data": {
              "chartID": "chart" + chartID,
              "chartType": "pie3DChart",
              "chartData": this.extractChartData(plotArea[key]["c:ser"])
            }
          };
          break;
        case "c:areaChart":
          chartData = {
            "type": "createChart",
            "data": {
              "chartID": "chart" + chartID,
              "chartType": "areaChart",
              "chartData": this.extractChartData(plotArea[key]["c:ser"])
            }
          };
          break;
        case "c:scatterChart":
          chartData = {
            "type": "createChart",
            "data": {
              "chartID": "chart" + chartID,
              "chartType": "scatterChart",
              "chartData": this.extractChartData(plotArea[key]["c:ser"])
            }
          };
          break;
        case "c:catAx":
          break;
        case "c:valAx":
          break;
        default:
      }
    }

    return result;
  }

  extractChartData(serNode: any) {
    var dataMat = new Array();

    if (serNode === undefined) {
      return dataMat;
    }

    if (serNode["c:xVal"] !== undefined) {
      var dataRow = new Array();
      this.eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode: any, index: number) => {
        dataRow.push(parseFloat(innerNode["c:v"]));
        return "";
      });
      dataMat.push(dataRow);
      dataRow = new Array();
      this.eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode: any, index: number) => {
        dataRow.push(parseFloat(innerNode["c:v"]));
        return "";
      });
      dataMat.push(dataRow);
    } else {
      this.eachElement(serNode, (innerNode: any, index: number) => {
        var dataRow = new Array();
        var colName = extractTextByPath(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

        // Category (string or number)
        let rowNames: any = {};
        if (extractTextByPath(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
          this.eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], (innerNode: any, index: number) => {
            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
            return "";
          });
        } else if (extractTextByPath(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
          this.eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], (innerNode: any, index: number) =>{
            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
            return "";
          });
        }

        // Value
        if (extractTextByPath(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
          this.eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"],  (innerNode: any, index: number) => {
            dataRow.push({ x: innerNode["attrs"]["idx"], y: parseFloat(innerNode["c:v"]) });
            return "";
          });
        }

        dataMat.push({ key: colName, values: dataRow, xlabels: rowNames });
        return "";
      });
    }

    return dataMat;
  }

  eachElement(node: any, doFunction: any) {
    if (node === undefined) {
        return;
    }
    var result = "";
    if (node.constructor === Array) {
        var l = node.length;
        for (var i=0; i<l; i++) {
            result += doFunction(node[i], i);
        }
    } else {
        result += doFunction(node, 0);
    }
    return result;
}

  genDiagram(node: any) {
    var order = node["attrs"]["order"];
    var xfrmNode = extractTextByPath(node, ["p:xfrm"]);
    return "<div class='block content' style='border: 1px dotted;" +
      this.getPosition(xfrmNode, undefined, undefined) + this.getSize(xfrmNode, undefined, undefined) +
      "'>TODO: diagram</div>";
  }
}