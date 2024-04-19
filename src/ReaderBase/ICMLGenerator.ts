import { ITranslator } from "./ITranslator";

export interface ICMLGenerator {
  generateCML: (
    objTranslator: ITranslator,
    FilePath: string,
    strDocumentName: string
  ) => Promise<XMLDocument>;
}
