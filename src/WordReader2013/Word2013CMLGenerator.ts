import { ICMLGenerator } from "../ReaderBase/ICMLGenerator";
import { ITranslator } from "../ReaderBase/ITranslator";
import * as xmldom from "@xmldom/xmldom";
import { IWordTranslator } from "../ReaderBase/IWordTranslator";
import { Utils } from "./CommonClasses/Utils";

interface XmlDeclaration {
  version: string;
  encoding?: string; // Optional encoding attribute
  standalone?: string; // Optional standalone attribute
}

export class Word2013CMLGenerator implements ICMLGenerator {
  generateCML = async (
    objTranslator: ITranslator,
    FilePath: string,
    strDocumentName: string
  ) => {
    let CMLDocument: XMLDocument =
      new xmldom.DOMImplementation().createDocument(null, null, null); // XmlDocument CMLDocument = new XmlDocument();
    const objWordTranslator = objTranslator as IWordTranslator;
    let xmlWordDoc: Element;
    let xmlBody: Element;

    const parsed = await objWordTranslator.ParseDocument(
      FilePath,
      CMLDocument,
      strDocumentName
    );
    if (parsed) {
      const xmldecl: ProcessingInstruction =
        CMLDocument.createProcessingInstruction(
          "xml",
          'version="1.0" encoding="UTF-8"'
        );
      CMLDocument.appendChild(xmldecl);

      // Create the root element 'cmlwddoc'
      xmlWordDoc = CMLDocument.createElement("cmlwddoc");
      xmlWordDoc.appendChild(CMLDocument.createTextNode(""));
      CMLDocument.appendChild(xmlWordDoc);

      const xmlTemp = await objWordTranslator.ReturnWordDocumentProperties();

      if (xmlTemp != null) {
        xmlWordDoc.appendChild(xmlTemp);
      }

      xmlBody = Utils.CreateNode(CMLDocument, xmlWordDoc, "body", "");
      Utils.CreateNode(CMLDocument, xmlBody, "props", "");
      const nBodyChiildCount: number = objWordTranslator.ReturnBodyChildCount();

      for (let i = 0; i < nBodyChiildCount; i++) {
        const xmlTemp = await objWordTranslator.ReturnBodyChild(i);
        if (xmlTemp != null) {
          xmlBody.appendChild(xmlTemp);
        }
      }

      return CMLDocument;
    } else {
      throw new Error("Document was not parsed correctly");
    }

    // return CMLDocument;
  };
}
