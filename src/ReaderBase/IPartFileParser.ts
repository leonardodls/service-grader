import { PackageReader } from "./PackageReader";

export interface IPartFileParser {
  retrunParsedElement: (
    partFileStream: Buffer,
    cmlDocumentXMLNode: XMLDocument,
    currentPackageReader: PackageReader
  ) => XMLDocument;
}
