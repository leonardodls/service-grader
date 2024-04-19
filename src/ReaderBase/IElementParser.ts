import { PackageReader } from "./PackageReader";

export interface IElementParser {
  ReturnParsedElement: (
    XmlOffice12Node: Element,
    CMLDocumentXMLNode: XMLDocument,
    CurrentPackageReader: PackageReader
  ) => Promise<Element>;
}
