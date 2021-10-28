export default class XMLTransform {
  _order = 1

  toXML(S: string) {
    let openBracket = "<";
    let openBracketCC = "<".charCodeAt(0);
    let closeBracket = ">";
    let closeBracketCC = ">".charCodeAt(0);
    let minus = "-";
    let minusCC = "-".charCodeAt(0);
    let slash = "/";
    let slashCC = "/".charCodeAt(0);
    let exclamation = '!';
    let exclamationCC = '!'.charCodeAt(0);
    let singleQuote = "'";
    let singleQuoteCC = "'".charCodeAt(0);
    let doubleQuote = '"';
    let doubleQuoteCC = '"'.charCodeAt(0);
    let questionMark = '?';
    let questionMarkCC = '?'.charCodeAt(0);

    /**
     *    returns text until the first nonAlphebetic letter
     */
    let nameSpacer = "\r\n\t>/= ";
    let pos = 0;
    function parseChildren(): any[] {
      var children = [];
      while (S[pos]) {
        if (S.charCodeAt(pos) == openBracketCC) {
          if (S.charCodeAt(pos + 1) === slashCC) { // </
            //while (S[pos]!=='>') { pos++; }
            pos = S.indexOf(closeBracket, pos);
            return children;
          } else if (S.charCodeAt(pos + 1) === exclamationCC) { // <! or <!--
            if (S.charCodeAt(pos + 2) == minusCC) {
              // comment support
              while (!(S.charCodeAt(pos) === closeBracketCC && S.charCodeAt(pos - 1) == minusCC &&
                S.charCodeAt(pos - 2) == minusCC && pos != -1)) {
                pos = S.indexOf(closeBracket, pos + 1);
              }
              if (pos === -1) {
                pos = S.length;
              }
            } else {
              // doctype support
              pos += 2;
              for (; S.charCodeAt(pos) !== closeBracketCC; pos++) { }
            }
            pos++;
            continue;
          } else if (S.charCodeAt(pos + 1) === questionMarkCC) { // <?
            // XML header support
            pos = S.indexOf(closeBracket, pos);
            pos++;
            continue;
          }
          pos++;
          var startNamePos = pos;
          for (; nameSpacer.indexOf(S[pos]) === -1; pos++) { }
          var node_tagName = S.slice(startNamePos, pos);

          // Parsing attributes
          var attrFound = false;
          var node_attributes: any = {};
          for (; S.charCodeAt(pos) !== closeBracketCC; pos++) {
            var c = S.charCodeAt(pos);
            if ((c > 64 && c < 91) || (c > 96 && c < 123)) {
              startNamePos = pos;
              for (; nameSpacer.indexOf(S[pos]) === -1; pos++) { }
              var name = S.slice(startNamePos, pos);
              // search beginning of the string
              var code = S.charCodeAt(pos);
              while (code !== singleQuoteCC && code !== doubleQuoteCC) {
                pos++;
                code = S.charCodeAt(pos);
              }

              var startChar = S[pos];
              var startStringPos = ++pos;
              pos = S.indexOf(startChar, startStringPos);
              var value = S.slice(startStringPos, pos);
              if (!attrFound) {
                node_attributes = {};
                attrFound = true;
              }
              node_attributes[name] = value;
            }
          }

          let node_children: any[] = []

          // Optional parsing of children
          if (S.charCodeAt(pos - 1) !== slashCC) {
            pos++;
            node_children = parseChildren();
          }

          children.push({
            "children": node_children,
            "tagName": node_tagName,
            "attrs": node_attributes
          });

        } else {
          var startTextPos = pos;
          pos = S.indexOf(openBracket, pos) - 1; // Skip characters until '<'
          if (pos === -2) {
            pos = S.length;
          }
          var text = S.slice(startTextPos, pos + 1);
          if (text.trim().length > 0) {
            children.push(text);
          }
        }
        pos++;
      }
      return children;
    }

    this._order = 1;
    return this.simplefy(parseChildren());
  }

  simplefy(children: any[]) {
    var node: any = {};

    if (children === undefined) {
      return {};
    }

    // Text node (e.g. <t>This is text.</t>)
    // if (children.length === 1 && typeof children[0] == 'string') {
    //     return children[0];
    // }

    // map each object
    let that = this
    children.forEach((child: any) => {
      if (!node[child.tagName]) {
        node[child.tagName] = [];
      }

      if (typeof child === 'object') {
        var kids = this.simplefy(child.children);
        if (child.attrs) {
          kids.attrs = child.attrs;
        }

        if (kids["attrs"] === undefined) {
          kids["attrs"] = { "order": this._order };
        } else {
          kids["attrs"]["order"] = this._order;
        }
        this._order++;
        node[child.tagName].push(kids);
      }
    });

    for (var i in node) {
      if (node[i].length == 1) {
        node[i] = node[i][0];
      }
    }

    return node;
  };
}
