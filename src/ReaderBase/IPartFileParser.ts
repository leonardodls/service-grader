export interface IPartFileParser {
  RetrunParsedElement: (
    partFileStream: Buffer,
    CMLDocumentXMLNode: XMLDocument
  ) => XMLDocument | HTMLElement;
}
