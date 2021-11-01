import { NodeElement } from "./model";

export interface Drawer {
  draw(nodes: NodeElement[]): string
}

export class HtmlDrawer implements Drawer {

  draw(nodes: NodeElement[]): string {
    for (const node of nodes) {
      console.log(node.eleType)
    }

    return ""
  }
}
