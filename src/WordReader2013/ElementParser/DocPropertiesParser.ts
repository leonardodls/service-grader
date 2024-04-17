import { IPartFileParser } from "../../ReaderBase/IPartFileParser";
import { Utils } from "../CommonClasses/Utils";
import xmldom from "xmldom";

export class DocPropertiesParser implements IPartFileParser {
  RetrunParsedElement = (
    partFileStream: Buffer,
    CMLDocumentXMLNode: XMLDocument
  ) => {
    const parser = new xmldom.DOMParser();
    const xmlTempDoc: XMLDocument = parser.parseFromString(
      partFileStream.toString(),
      "text/xml"
    );

    // const docProps: DocPropertiesClass = new DocPropertiesClass("", "props", "", CMLDocumentXMLNode); - need to determine better approach

    const docProps = CMLDocumentXMLNode.createElement("props");

    const xmlEle: Element = xmlTempDoc.lastChild as Element;

    if (xmlEle != null && xmlEle.localName == "Properties") {
      let temp: Element = xmlEle.firstChild as Element;
      while (temp != null) {
        Utils.createNode(
          CMLDocumentXMLNode,
          docProps,
          temp.localName,
          String(temp.textContent)
        );
        temp = temp.nextSibling as Element;
      }
    }
    return docProps;
  };
}
