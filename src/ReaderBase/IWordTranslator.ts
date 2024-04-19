import { ITranslator } from "./ITranslator";

export interface IWordTranslator extends ITranslator {
  ReturnWordDocumentProperties: () => Promise<XMLDocument | Element | null>;
  ReturnCoreProperties: (eleDocProps: XMLDocument) => XMLDocument | Element; //need to add in interface as virtual functions not supported in typescript.
  ReturnBodyChildCount: () => number; //need to add in interface as virtual functions not supported in typescript.
  ReturnBodyChild: (nChildNo: number) => Promise<Element | null>; //need to add in interface as virtual functions not supported in typescript.
}
