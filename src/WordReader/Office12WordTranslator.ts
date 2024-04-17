import { IPartFileParser } from "../ReaderBase/IPartFileParser";
import { IWordTranslator } from "../ReaderBase/IWordTranslator";
import { PackageReader } from "../ReaderBase/PackageReader";
import { CommonFunctions } from "../Utils/common";
import { DocPropertiesParser } from "../WordReader2013/ElementParser/DocPropertiesParser";
import * as fs from "node:fs/promises";

export class Office12WordTranslator implements IWordTranslator {
  protected m_myPackageReader: PackageReader | null = null;
  protected m_DocumentName: string = "";
  protected m_xmlGeneDoc: XMLDocument | null = null;

  parseDocument = async (
    filePath: string,
    cmlDocumentXMLNode: XMLDocument,
    strDocumentName: string
  ) => {
    this.m_DocumentName = strDocumentName;
    // Utils utils = new Utils();
    const docReader = await fs.readFile(filePath, null);

    let strEtension: string = "";
    if (this.m_DocumentName != "" && this.m_DocumentName.length > 4) {
      strEtension = this.m_DocumentName.substring(
        this.m_DocumentName.length - 4,
        4
      );
    }

    if (this.m_DocumentName == "" || strEtension.toLowerCase() != "docx") {
      throw new Error("The package does not contain the basic document.xml");
    }

    this.m_xmlGeneDoc = cmlDocumentXMLNode;
    this.m_myPackageReader = new PackageReader();

    const checkForIncompatibleOfficePlatform: boolean = true;
    this.m_myPackageReader.initalisePackage(
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

    const uri = new URL(CommonFunctions.prependStringToURIPath(srtURI, "/"));

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

  ReturnCoreProperties = (eleDocProps: XMLDocument): XMLDocument => {
    if (!this.m_myPackageReader) {
      throw new Error("Package is not initailized yet..");
    }
    // PackageProperties pp = this.m_myPackageReader.ReturnPackageProperties();
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "category", pp.Category);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "contentstatus", pp.ContentStatus);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "contenttype", pp.ContentType);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "creatd", pp.Created.ToString());
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "creator", pp.Creator);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "desc", pp.Description);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "identifier", pp.Identifier);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "keywords", pp.Keywords);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "lang", pp.Language);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "lstmodby", pp.LastModifiedBy);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "lstprntd", pp.LastPrinted.ToString());
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "modfd", pp.Modified.ToString());
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "revision", pp.Revision);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "subject", pp.Subject);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "title", pp.Title);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "ver", pp.Version);
    // //windows theme clr
    // string sAuto = GetThemeClrCode();
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "autoclr", sAuto);
    return eleDocProps;
  };
}
