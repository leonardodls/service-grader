export interface ITranslator {
  ParseDocument: (
    FilePath: string,
    CMLDocumentXMLNode: XMLDocument,
    strDocumentName: string
  ) => Promise<boolean>;
}
