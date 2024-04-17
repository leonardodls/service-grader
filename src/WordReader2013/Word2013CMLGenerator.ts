import { ICMLGenerator } from "../ReaderBase/ICMLGenerator";
import { ITranslator } from "../ReaderBase/ITranslator";
import * as xmldom from "xmldom";
import { IWordTranslator } from "../ReaderBase/IWordTranslator";

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
    let cmlDocument: XMLDocument =
      new xmldom.DOMImplementation().createDocument(null, null, null);
    const objWordTranslator = objTranslator as IWordTranslator;
    let xmlWordDoc: HTMLElement;

    const parsed = await objWordTranslator.ParseDocument(
      FilePath,
      cmlDocument,
      strDocumentName
    );
    if (parsed) {
      const xmldecl: ProcessingInstruction =
        cmlDocument.createProcessingInstruction(
          "xml",
          'version="1.0" encoding="UTF-8"'
        );
      cmlDocument.appendChild(xmldecl);

      // Create the root element 'cmlwddoc'
      xmlWordDoc = cmlDocument.createElement("cmlwddoc");
      xmlWordDoc.appendChild(cmlDocument.createTextNode(""));
      cmlDocument.appendChild(xmlWordDoc);

      const xmlTemp = await objWordTranslator.ReturnWordDocumentProperties();

      if (xmlTemp != null) {
        xmlWordDoc.appendChild(xmlTemp);
      }

      return cmlDocument;
    } else {
      throw new Error("Document was not parsed correctly");
    }

    // return CMLDocument;
  };
}
