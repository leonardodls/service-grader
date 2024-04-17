export abstract class CMLBaseElement extends XMLDocument {
  constructor(s1: string, s2: string, s3: string, doc: XMLDocument) {
    super();
  }

  abstract returnElementType(): string;
}
