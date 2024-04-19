import { CMLBaseElement } from "../../ReaderBase/BaseElement";

export class DocPropertiesClass extends CMLBaseElement {
  //protected internal XmlElement (string? prefix, string localName, string? namespaceURI, System.Xml.XmlDocument doc);
  constructor(s1: string, s2: string, s3: string, doc: XMLDocument) {
    super(s1, s2, s3, doc);
  }

  returnElementType() {
    return this.toString();
  }
}
