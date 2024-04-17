import { ITranslator } from "./ITranslator";

export interface ICMLGenerator {
  generateCML: (
    objTranslator: ITranslator,
    filePath: string,
    strDocumentName: string
  ) => Promise<XMLDocument>;
}
