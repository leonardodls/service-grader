import { IPartFileParser } from "../ReaderBase/IPartFileParser";
import { IWordTranslator } from "../ReaderBase/IWordTranslator";
import { PackageReader } from "../ReaderBase/PackageReader";
import { CommonFunctions } from "../Utils/common";
import { Utils } from "../WordReader2013/CommonClasses/Utils";
import { DocPropertiesParser } from "../WordReader2013/ElementParser/DocPropertiesParser";
import * as fs from "node:fs/promises";
import xpath from "xpath";

export class Office12WordTranslator implements IWordTranslator {
  protected m_myPackageReader: PackageReader | null = null;
  protected m_DocumentName: string = "";
  protected m_xmlGeneDoc: XMLDocument | null = null;
  protected m_ThemeDoc: XMLDocument | null = null;
  protected m_xmlDoc: XMLDocument | null = null;
  protected m_xmlBodyEle: Element | null = null;
  protected m_Uri: string = "";
  m_nChildIndex: number = 0;
  public m_StylesDoc: XMLDocument | null = null;

  ParseDocument = async (
    FilePath: string,
    CMLDocumentXMLNode: XMLDocument,
    strDocumentName: string
  ) => {
    this.m_DocumentName = strDocumentName;
    // Utils utils = new Utils();
    const docReader = await fs.readFile(FilePath, null);

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

    this.m_xmlGeneDoc = CMLDocumentXMLNode;
    this.m_myPackageReader = new PackageReader();

    const checkForIncompatibleOfficePlatform: boolean = true;
    this.m_myPackageReader.initalisePackage(
      docReader,
      checkForIncompatibleOfficePlatform
    );

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

    const uri = CommonFunctions.PrependStringToURIPath(srtURI, "/"); //Uri uri = new Uri(CommonFunctions.PrependStringToURIPath(srtURI, "/"), UriKind.Relative);
    const ss: Buffer | null = await this.m_myPackageReader.ReturnPackagePart(
      uri
    ); // Stream ss = m_myPackageReader.ReturnPackagePart(uri);
    const docProps: IPartFileParser = new DocPropertiesParser();
    const eleDocProps: XMLDocument | Element = docProps.RetrunParsedElement(
      ss as Buffer,
      this.m_xmlGeneDoc as XMLDocument
    );

    const xReturn = this.ReturnCoreProperties(eleDocProps);

    return xReturn;
  };

  ReturnCoreProperties = (
    eleDocProps: XMLDocument | Element
  ): XMLDocument | Element => {
    if (!this.m_myPackageReader) {
      throw new Error("Package is not initailized yet..");
    }
    // PackageProperties pp = this.m_myPackageReader.ReturnPackageProperties();
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "category", pp.Category);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "contentstatus", pp.ContentStatus);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "contenttype", pp.ContentType);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "creatd", pp.Created.toString());
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "creator", pp.Creator);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "desc", pp.Description);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "identifier", pp.Identifier);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "keywords", pp.Keywords);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "lang", pp.Language);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "lstmodby", pp.LastModifiedBy);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "lstprntd", pp.LastPrinted.toString());
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "modfd", pp.Modified.toString());
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "revision", pp.Revision);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "subject", pp.Subject);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "title", pp.Title);
    // Utils.CreateNode(m_xmlGeneDoc, eleDocProps, "ver", pp.Version);
    // //windows theme clr
    const sAuto = this.GetThemeClrCode();
    Utils.CreateNode(
      this.m_xmlGeneDoc as XMLDocument,
      eleDocProps,
      "autoclr",
      sAuto
    );
    return eleDocProps;
  };

  ReturnBodyChildCount = () => {
    if (!this.m_xmlDoc) {
      throw new Error("Package is not parsed yet..");
    }
    const tempList: HTMLCollectionOf<Element> =
      this.m_xmlDoc.getElementsByTagName("w:body");

    this.m_xmlBodyEle = tempList.item(0);
    if (tempList.length != 0 && this.m_xmlBodyEle) {
      return this.m_xmlBodyEle.childNodes.length - 1;
    }
    return 0;
  };

  ReturnBodyChild = async (nChildNo: number): Promise<Element | null> => {
    return null;
  };

  LoadXMLs = async () => {
    if (!this.m_myPackageReader) {
      throw new Error("Package is not parsed yet..");
    }
    //Styles.xml
    let sXml = await this.m_myPackageReader.ReturnBaseXML(
      this.m_Uri,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    );
    let sPartFileName = CommonFunctions.PrependStringToURIPath(sXml, "/word/");
    this.m_StylesDoc = await this.m_myPackageReader.ReturnDocmentFromPart(
      sPartFileName
    );
    this.m_myPackageReader.partFileMap.set("m_StylesDoc", this.m_StylesDoc);
    //Themes.xml
    sXml = await this.m_myPackageReader.ReturnBaseXML(
      this.m_Uri,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    );

    if (sXml != "") {
      const sPartFileName = CommonFunctions.PrependStringToURIPath(
        sXml,
        "/word/"
      );
      this.m_ThemeDoc = await this.m_myPackageReader.ReturnDocmentFromPart(
        sPartFileName
      );
      this.m_myPackageReader.partFileMap.set("m_ThemeDoc", this.m_ThemeDoc);
    }
  };

  protected GetThemeClrCode(): string {
    let strXPath = "/a:theme/a:themeElements/a:clrScheme/a:dk1/a:sysClr";
    const select = xpath.useNamespaces({
      a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    });

    // Explicitly cast the result to an array of Element
    let xElements = select(
      strXPath,
      this.m_ThemeDoc as XMLDocument
    ) as Element[];

    if (xElements.length > 0 && xElements[0]) {
      let xElement = xElements[0];
      if (xElement.getAttribute("val") === "windowText") {
        return xElement.getAttribute("lastClr") || "";
      }
    } else {
      strXPath = "/a:theme/a:themeElements/a:clrScheme/a:dk1/a:srgbClr";
      xElements = select(strXPath, this.m_ThemeDoc as XMLDocument) as Element[];
      if (
        xElements.length > 0 &&
        xElements[0] &&
        xElements[0].hasAttribute("val")
      ) {
        return xElements[0].getAttribute("val") || "";
      }
    }
    return "";
  }
}
