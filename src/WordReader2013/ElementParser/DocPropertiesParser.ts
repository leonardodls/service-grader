import { IPartFileParser } from "../../ReaderBase/IPartFileParser";
import { Utils } from "../CommonClasses/Utils";
import xmldom from "@xmldom/xmldom";

export class DocPropertiesParser implements IPartFileParser {
  RetrunParsedElement = (
    partFileStream: Buffer,
    CMLDocumentXMLNode: XMLDocument
  ) => {
    const parser = new xmldom.DOMParser();
    const xmlTempDoc: XMLDocument = parser.parseFromString(
      partFileStream.toString(),
      "text/xml"
    ); // XmlDocument xmlTempDoc = new XmlDocument(); + xmlTempDoc.Load(PartFileStream);

    const docProps = CMLDocumentXMLNode.createElement("props"); // const docProps: DocPropertiesClass = new DocPropertiesClass("", "props", "", CMLDocumentXMLNode); - need to determine better approach

    const xmlEle: Element = xmlTempDoc.lastChild as Element;

    if (xmlEle != null && xmlEle.localName == "Properties") {
      let temp: Element = xmlEle.firstChild as Element;
      while (temp != null) {
        Utils.CreateNode(
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
