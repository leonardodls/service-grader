import { IElementParser } from "../../ReaderBase/IElementParser";
import { PackageReader } from "../../ReaderBase/PackageReader";
import { CommonFunctions } from "../../Utils/common";
import { TextAttributes2013 } from "../../WordReader2013/CommonClasses/TextAttributes";
import { Utils } from "../../WordReader2013/CommonClasses/Utils";

export class ParagraphParser implements IElementParser {
  protected m_xmlParaEle: Element | null = null;
  protected m_paragraph: Element | null = null;
  protected m_myPackageReader: PackageReader | null = null;
  protected m_myXmlDoc: XMLDocument | null = null;
  protected m_docUri: string = "";

  ReturnParsedElement = async (
    XmlOffice12Node: Element,
    CMLDocumentXMLNode: XMLDocument,
    CurrentPackageReader: PackageReader
  ) => {
    this.m_xmlParaEle = XmlOffice12Node;
    this.m_myPackageReader = CurrentPackageReader;
    this.m_myXmlDoc = CMLDocumentXMLNode;

    // Utils.InitailizeNamespace(ref m_nsmgr, m_myXmlDoc, "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

    const strDocUriString: string = await this.m_myPackageReader.ReturnBaseXML(
      null,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    );
    this.m_docUri = CommonFunctions.PrependStringToURIPath(
      strDocUriString,
      "/"
    );

    this.m_paragraph = this.m_myXmlDoc.createElement("p"); // this.m_paragraph = new ParagraphClass("", "p", "", m_myXmlDoc);

    const pProps: Element = Utils.CreateNode(
      this.m_myXmlDoc,
      this.m_paragraph,
      "props",
      ""
    );

    const txtAtt: TextAttributes2013 = new TextAttributes2013(); //TextAttributes2013 txtAtt = new TextAttributes2013();
    txtAtt.m_bParaInCell = false; //txtAtt.m_bParaInCell = this.m_bParaInCell;
    txtAtt.ParseTextAttributes(
      XmlOffice12Node,
      this.m_myXmlDoc,
      pProps,
      this.m_myPackageReader
    );
    txtAtt.GenerateParaText(XmlOffice12Node, pProps);
    return this.m_paragraph;
  };
}
