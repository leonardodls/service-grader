export class Utils {
  static CreateNode = (
    xmlDoc: XMLDocument,
    xmlParent: XMLDocument | Element,
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

  static CreateAttribute = (
    xmlDoc: XMLDocument,
    xmlParent: Element,
    attrName: string,
    attrValue: string
  ) => {
    const xmlAttr = xmlDoc.createAttribute(attrName);
    xmlAttr.value = attrValue;
    xmlParent.attributes.setNamedItem(xmlAttr);
    return xmlAttr;
  };
}
