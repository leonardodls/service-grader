import * as fs from "node:fs/promises";
import { PackageReader } from "../ReaderBase/PackageReader";
import { CommonFunctions } from "../Utils/common";
import { IPartFileParser } from "../ReaderBase/IPartFileParser";
import { DocPropertiesParser } from "./ElementParser/DocPropertiesParser";
import { Office12WordTranslator } from "../WordReader/Office12WordTranslator";
import { Utils } from "./CommonClasses/Utils";
import xmldom from "@xmldom/xmldom";
import { ParagraphParser2013 } from "./ElementClasses/ParagraphParser2013";

export class Word2013Translator extends Office12WordTranslator {
  ParseDocument = async (
    FilePath: string,
    CMLDocumentXMLNode: XMLDocument,
    strDocumentName: string
  ): Promise<boolean> => {
    this.m_DocumentName = strDocumentName;

    const docReader = await fs.readFile(FilePath, null);

    let strEtension: string = "";
    if (this.m_DocumentName != "" && this.m_DocumentName.length > 4) {
      strEtension = this.m_DocumentName.substring(
        this.m_DocumentName.length - 4
      );
    }

    if (this.m_DocumentName == "" || strEtension.toLowerCase() != "docx") {
      throw new Error("The package does not contain the basic document.xml");
    }

    this.m_xmlGeneDoc = CMLDocumentXMLNode;
    this.m_myPackageReader = new PackageReader();

    const checkForIncompatibleOfficePlatform: boolean = true;
    await this.m_myPackageReader.initalisePackage(
      docReader,
      checkForIncompatibleOfficePlatform
    );

    const uriString: string = await this.m_myPackageReader.ReturnBaseXML(
      null,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    );

    if (uriString == "") {
      return false;
    }
    this.m_Uri = CommonFunctions.PrependStringToURIPath(uriString, "/");

    const ss: Buffer | null = await this.m_myPackageReader.ReturnPackagePart(
      this.m_Uri,
      0
    );

    if (!ss) {
      return false;
    }

    //Read the document.xml
    const parser = new xmldom.DOMParser();
    this.m_xmlDoc = parser.parseFromString(ss.toString(), "text/xml");

    this.LoadXMLs();

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

    // const uri = new URL(CommonFunctions.PrependStringToURIPath(srtURI, "/")); need to understand better
    const uri = CommonFunctions.PrependStringToURIPath(srtURI, "/");

    const ss: Buffer | null = await this.m_myPackageReader.ReturnPackagePart(
      uri
    );
    const docProps: IPartFileParser = new DocPropertiesParser();

    const eleDocProps: XMLDocument | Element = docProps.RetrunParsedElement(
      ss as Buffer,
      this.m_xmlGeneDoc as XMLDocument
    );

    Utils.CreateNode(
      this.m_xmlGeneDoc as XMLDocument,
      eleDocProps,
      "name",
      this.m_DocumentName
    );

    const xReturn = this.ReturnCoreProperties(eleDocProps);

    return xReturn;
  };

  ReturnBodyChild = async (nChildNo: number) => {
    if (!this.m_xmlBodyEle) {
      throw new Error("Package is not parsed yet..");
    }
    const xmlChildEle = this.m_xmlBodyEle.childNodes.item(nChildNo) as Element;

    switch (xmlChildEle.nodeName) {
      case "w:p":
        let xPara: Element;
        try {
          const paragraph: ParagraphParser2013 = new ParagraphParser2013();
          xPara = await paragraph.ReturnParsedElement(
            xmlChildEle,
            this.m_xmlGeneDoc as XMLDocument,
            this.m_myPackageReader as PackageReader
          );
        } catch (error) {
          throw new Error(
            "Something wen wrong in ReturnBodyChild method of word2013 translator!!"
          );
        }
        Utils.CreateAttribute(
          this.m_xmlGeneDoc as XMLDocument,
          xPara,
          "index",
          (++this.m_nChildIndex).toString()
        );

        return xPara;

      default:
        return null;
    }
  };
}
