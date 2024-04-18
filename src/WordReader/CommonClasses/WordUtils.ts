import { Utils } from "../../WordReader2013/CommonClasses/Utils";

export class WordUtils {
  public static ReturnNode = (
    xmlDoc: XMLDocument,
    xmlParent: Element,
    nodeToBeChecked: string
  ) => {
    let xmlReturnEle = null;
    xmlReturnEle = xmlParent.getElementsByTagName(nodeToBeChecked)[0];

    if (xmlReturnEle == null) {
      xmlReturnEle = Utils.CreateNode(xmlDoc, xmlParent, nodeToBeChecked, "");
    }
    return xmlReturnEle;
  };
}
