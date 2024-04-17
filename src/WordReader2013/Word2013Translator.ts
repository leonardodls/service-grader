import * as fs from "node:fs/promises";
import { PackageReader } from "../ReaderBase/PackageReader";
import { CommonFunctions } from "../Utils/common";
import { IPartFileParser } from "../ReaderBase/IPartFileParser";
import { DocPropertiesParser } from "./ElementParser/DocPropertiesParser";
import { Office12WordTranslator } from "../WordReader/Office12WordTranslator";

export class Word2013Translator extends Office12WordTranslator {
  parseDocument = async (
    filePath: string,
    cmlDocumentXMLNode: XMLDocument,
    strDocumentName: string
  ): Promise<boolean> => {
    this.m_DocumentName = strDocumentName;
    // Utils utils = new Utils();
    const docReader = await fs.readFile(filePath, null);

    let strEtension: string = "";
    if (this.m_DocumentName != "" && this.m_DocumentName.length > 4) {
      strEtension = this.m_DocumentName.substring(
        this.m_DocumentName.length - 4
      );
    }

    if (this.m_DocumentName == "" || strEtension.toLowerCase() != "docx") {
      throw new Error("The package does not contain the basic document.xml");
    }

    this.m_xmlGeneDoc = cmlDocumentXMLNode;
    this.m_myPackageReader = new PackageReader();

    const checkForIncompatibleOfficePlatform: boolean = true;
    await this.m_myPackageReader.initalisePackage(
      docReader,
      checkForIncompatibleOfficePlatform
    );

    return true;
  };

  ReturnWordDocumentProperties = async () => {
    if (!this.m_myPackageReader) {
      throw new Error("Package is not initailized yet..");
    }
    const srtURI: string = await this.m_myPackageReader.ReturnBaseXML(
      null,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    );
    if (srtURI == "") {
      return null;
    }

    // const uri = new URL(CommonFunctions.prependStringToURIPath(srtURI, "/")); need to understand better
    const uri = CommonFunctions.prependStringToURIPath(srtURI, "/");

    const ss: Buffer | null = await this.m_myPackageReader.ReturnPackagePart(
      String(uri)
    );
    const docProps: IPartFileParser = new DocPropertiesParser();

    const eleDocProps: XMLDocument = docProps.retrunParsedElement(
      ss as Buffer,
      this.m_xmlGeneDoc as XMLDocument,
      this.m_myPackageReader
    );

    const xReturn = this.ReturnCoreProperties(eleDocProps);

    return xReturn;
  };
}
