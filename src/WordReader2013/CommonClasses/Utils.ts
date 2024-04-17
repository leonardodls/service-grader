export class Utils {
  static createNode = (
    xmlDoc: XMLDocument,
    xmlParent: XMLDocument,
    tagName: string,
    tagValue: string
  ): Element => {
    tagName = tagName.toLowerCase();
    let xmlEle = null;

    xmlEle = xmlDoc.createElement(tagName);

    xmlEle.appendChild(xmlDoc.createTextNode(tagValue));
    xmlParent.appendChild(xmlEle);
    return xmlEle;
  };
}
