import { IPartFileParser } from "../../ReaderBase/IPartFileParser";
import { PackageReader } from "../../ReaderBase/PackageReader";
import { Utils } from "../CommonClasses/Utils";
import { DocPropertiesClass } from "../ElementClasses/DocPropertiesClass";
import { DOMImplementation, DOMParser } from "xmldom";

export class DocPropertiesParser implements IPartFileParser {
  retrunParsedElement = (
    partFileStream: Buffer,
    cmlDocumentXMLNode: XMLDocument,
    currentPackageReader: PackageReader
  ) => {
    const parser = new DOMParser();
    const xmlTempDoc: XMLDocument = parser.parseFromString(
      partFileStream.toString(),
      "text/xml"
    );

    // const docProps: DocPropertiesClass = new DocPropertiesClass("", "props", "", cmlDocumentXMLNode); - need to determine better approach
    const docProps = new DOMImplementation().createDocument(null, null, null);
    cmlDocumentXMLNode.appendChild(docProps.createElement("props"));

    const xmlEle: Element = xmlTempDoc.lastChild as Element;

    if (xmlEle != null && xmlEle.localName == "Properties") {
      let temp: Element = xmlEle.firstChild as Element;
      while (temp != null) {
        Utils.createNode(
          cmlDocumentXMLNode,
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
