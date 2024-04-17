export interface ITranslator {
  parseDocument: (
    filePath: string,
    cmlDocumentXMLNode: XMLDocument,
    strDocumentName: string
  ) => Promise<boolean>;
}
