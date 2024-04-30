import { Utils } from "../../WordReader2013/CommonClasses/Utils";
import xpath from "xpath";
export class WordUtils {
  public static returnNode = (
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

  public static ConvertValues = (
    sConversionFactor: number,
    sToConvert: string,
    bRound: boolean
  ): string => {
    let sRetVal = "";
    let nTemp = parseFloat(sToConvert);

    nTemp /= sConversionFactor;
    if (bRound) {
      nTemp = parseFloat(nTemp.toFixed(2));
    }
    sRetVal = nTemp.toString();
    return sRetVal;
  };

  public static GetPatternStyle(sPattern: string): string {
    let sPatStyle = "";
    switch (sPattern) {
      case "solid":
        sPatStyle = "Solid(100%)";
        break;

      case "pct5":
        sPatStyle = "5%";
        break;

      case "pct10":
        sPatStyle = "10%";
        break;

      case "pct12":
        sPatStyle = "12.5%";
        break;

      case "pct15":
        sPatStyle = "15%";
        break;

      case "pct20":
        sPatStyle = "20%";
        break;

      case "pct25":
        sPatStyle = "25%";
        break;

      case "pct30":
        sPatStyle = "30%";
        break;

      case "pct35":
        sPatStyle = "35%";
        break;

      case "pct37":
        sPatStyle = "37.5%";
        break;

      case "pct40":
        sPatStyle = "40%";
        break;

      case "pct45":
        sPatStyle = "45%";
        break;

      case "pct50":
        sPatStyle = "50%";
        break;

      case "pct55":
        sPatStyle = "55%";
        break;

      case "pct60":
        sPatStyle = "60%";
        break;

      case "pct62":
        sPatStyle = "62.5%";
        break;

      case "pct65":
        sPatStyle = "65%";
        break;

      case "pct70":
        sPatStyle = "70%";
        break;

      case "pct75":
        sPatStyle = "75%";
        break;

      case "pct80":
        sPatStyle = "80%";
        break;

      case "pct85":
        sPatStyle = "85%";
        break;

      case "pct87":
        sPatStyle = "87.5%";
        break;

      case "pct90":
        sPatStyle = "90%";
        break;

      case "pct95":
        sPatStyle = "95%";
        break;

      case "horzStripe":
        sPatStyle = "Dk Horizontal";
        break;

      case "vertStripe":
        sPatStyle = "Dk Vertical";
        break;

      case "reverseDiagStripe":
        sPatStyle = "Dk Dwn Diagonal";
        break;

      case "diagStripe":
        sPatStyle = "Dk Up Diagonal";
        break;

      case "horzCross":
        sPatStyle = "Dk Grid";
        break;

      case "diagCross":
        sPatStyle = "Dk Trellis";
        break;

      case "thinHorzStripe":
        sPatStyle = "Lt Horizontal";
        break;

      case "thinVertStripe":
        sPatStyle = "Lt Vertical";
        break;

      case "thinReverseDiagStripe":
        sPatStyle = "Lt Dwn Diagonal";
        break;

      case "thinDiagStripe":
        sPatStyle = "Lt Up Diagonal";
        break;

      case "thinHorzCross":
        sPatStyle = "Lt Grid";
        break;

      case "thinDiagCross":
        sPatStyle = "Lt Trellis";
        break;
      default:
        sPatStyle = "Clear";
        break;
    }
    return sPatStyle;
  }

  public static extractStringFromVal(
    Office12Node: Element,
    strOffice12Node: string,
    select: xpath.XPathSelect
  ): string {
    if (Office12Node) {
      let strValue = "";
      const xmlTemp = select(strOffice12Node, Office12Node) as Element[];

      if (xmlTemp.length > 0) {
        const fontNode = xmlTemp[0].getAttribute("w:val");
        if (fontNode !== null) {
          strValue = fontNode;
        }
      }
      return strValue;
    } else {
      return "";
    }
  }

  public static extractString(
    Office12Node: Element,
    strOffice12Node: string,
    select: xpath.XPathSelect
  ) {
    let strValue = "";
    if (Office12Node) {
      const xmlTemp = select(strOffice12Node, Office12Node, true) as Element;
      if (xmlTemp != null) {
        strValue = "1";
      }
    }
    return strValue;
  }

  public static ReturnCnfIndexRef = (strCnfStyle: string): string => {
    let strIndexRef = "";
    const nIndex = strCnfStyle.indexOf("1");
    switch (nIndex) {
      case 0:
        strIndexRef = "firstRow";
        break;
      case 1:
        strIndexRef = "lastRow";
        break;
      case 2:
        strIndexRef = "firstCol";
        break;
      case 3:
        strIndexRef = "lastCol";
        break;
      case 4:
        strIndexRef = "band1Vert";
        break;
      case 5:
        strIndexRef = "band2Vert";
        break;
      case 6:
        strIndexRef = "band1Horz";
        break;
      case 7:
        strIndexRef = "band2Horz";
        break;
      case 8:
        strIndexRef = "neCell";
        break;
      case 9:
        strIndexRef = "nwCell";
        break;
      case 10:
        strIndexRef = "seCell";
        break;
      case 11:
        strIndexRef = "swCell";
        break;
      default:
        break;
    }
    return strIndexRef;
  };

  public static ExtractAttribute = (
    xElement: Element,
    sAttribute: string
  ): string => {
    return xElement.getAttribute(sAttribute) || "";
  };
}
