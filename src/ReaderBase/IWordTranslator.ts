import { ITranslator } from "./ITranslator";

export interface IWordTranslator extends ITranslator {
  ReturnWordDocumentProperties: () => Promise<XMLDocument | HTMLElement | null>;
  ReturnCoreProperties: (eleDocProps: XMLDocument) => XMLDocument | HTMLElement; //need to add in interface as virtual functions not supported in typescript.
}
