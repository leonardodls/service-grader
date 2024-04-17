import { ITranslator } from "./ITranslator";

export interface IWordTranslator extends ITranslator {
  ReturnWordDocumentProperties: () => Promise<XMLDocument | null>;
  ReturnCoreProperties: (eleDocProps: XMLDocument) => XMLDocument;
}
