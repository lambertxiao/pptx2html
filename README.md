
slide -> container
node -> element


interface Drawer {}
class HtmlDrawer implements Drawer {}


ZipLoder
XmlExplainer
NodeProcessor xml => node{textNode, shapeNode, graphic, img}
Drawer



[
  {
    bg: "",
    nodes: "",
  },
  {

  }
]

```
  <div class="container">
    <div class="row" style="width: {{width}}px;">
      <div class="col-md-24">
        <div id="myCarousel" class="carousel slide">
          <div class="carousel-inner">
            {{content}}
          </div>
          <a class="carousel-control left" href="#myCarousel" data-slide="prev"> 
            <span _ngcontent-c3="" aria-hidden="true" class="glyphicon glyphicon-chevron-left"></span>
          </a>
          <a class="carousel-control right" href="#myCarousel" data-slide="next">
            <span _ngcontent-c3="" aria-hidden="true" class="glyphicon glyphicon-chevron-right"></span>
          </a>
        </div>
      </div>
    </div>
  </div>
```