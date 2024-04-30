import { PackageReader } from "../../ReaderBase/PackageReader";
import xpath from "xpath";
import { CommonFunctions } from "../../Utils/common";
import { Utils } from "./Utils";
import { WordUtils } from "../../WordReader/CommonClasses/WordUtils";
import { WordUtils2010 } from "../../WordReader2010/CommonClasses/WordUtils";
import { ParaMarkAttributes2013 } from "./ParaMarkAttributes2013";

interface ColorAttribs {
  strClr: string;
  strTClr: string;
  strTShade: string;
  strTTint: string;
  strType: string;
}

interface TextOutlineAttribs {
  strThemColor: string;
  strThemeShade: string;
  strSatMode: string;
  strLineType: string;
  strWidth: string;
}

interface GradientStopAttributes {
  strPos: string;
  strTClr: string;
  strTSat: string;
  strTTint: string;
}

interface ShadeAttribs {
  strShdPattStyle: string;
  strShdPattCol: string;
  strShdPattThemeCol: string;
  strShdPattThemeShade: string;
  strShdPattThemeTint: string;
  strShdClr: string;
  strShdTClr: string;
  strShdTShade: string;
  strShdTTint: string;
}

interface UnderlineAttribs {
  strIsUnderline: string;
  strUnderlineClr: string;
  strUnderlineType: string;
  strUnderlineTClr: string;
  strUnderlineTShade: string;
  strUnderlineTTint: string;
}

interface ShadowAttribs {
  strShdowblur: string;
  strShdowdist: string;
  strShdowdir: string;
  strShdowsx: string;
  strShdowsy: string;
  strShdowSchemeclr: string;
  strShdowLummod: string;
  strShdowLumOff: string;
}

interface paraFontAttributes {
  strFontName: string;
  strFontSize: string;
  strType: string;
  strColor: string;
  strThemeColor: string;
  strThemeShade: string;
  strThemeTint: string;
  gradientStops: GradientStopAttributes[];

  strTextOutlineThemColor: string;
  strTextOutlineThemeShade: string;
  strTextOutlineSatMode: string;
  strTextOutlineLineType: string;
  strTextOutlineWidth: string;

  strShdColor: string;
  strShdThemeColor: string;
  strShdThemeShade: string;
  strShdThemeTint: string;
  strShdPattStyle: string;
  strShdPattCol: string;
  strShdPattThemeClr: string;
  strShdPattThemeShade: string;
  strShdPattThemeTint: string;

  strBold: string;
  strItalic: string;

  strIsUnderline: string;
  strUnderlineClr: string;
  strUnderlineType: string;
  strUnderlineClrThemeClr: string;
  strUnderlineClrThemeShade: string;
  strUnderlineClrThemeTint: string;

  strStrikethrough: string;
  strVAlign: string;
  strOutline: string;
  strEmboss: string;
  strEngrave: string;
  strShadow: string;
  strHidden: string;
  strCaps: string;
  strSmallCaps: string;
  strDoubleStrike: string;
  strHighlightColor: string;
  strSpacing: string;
  strWidth: string;
  strKerning: string;
  strPosition: string;
  strStyle: string;
  strFootnoteID: string;
  strEndnoteID: string;
  strBreakType: string;

  strPagination: string;

  strText: string;
  strAutoUpdate: string;
  strStyleType: string;

  strShdowblur: string;
  strShdowdist: string;
  strShdowdir: string;
  strShdowsx: string;
  strShdowsy: string;
  strShdowSchemeclr: string;
  strShdowLummod: string;
  strShdowLumOff: string;

  strLigature: string;
}

class ColorAttribs implements ColorAttribs {}
class GradientStopAttributes implements GradientStopAttributes {}
class ShadeAttribs implements ShadeAttribs {}
class UnderlineAttribs implements UnderlineAttribs {}
class TextOutlineAttribs implements TextOutlineAttribs {}
class ShadowAttribs implements ShadowAttribs {}
class paraFontAttributes implements paraFontAttributes {}

export class TextAttributes2013 {
  public m_bParaInCell: boolean = false;
  private m_xmlParaEle: Element | null = null;
  private nIndex: number = 0;
  private m_myXmlDoc: XMLDocument | null = null;
  private m_CMLProps: Element | null = null;
  private m_myPackageReader: PackageReader | null = null;
  protected m_strThemeMinorFont: string | null = "";
  protected m_strThemeMajorFont: string | null = "";
  protected m_docUri: string = "";
  protected m_paraAttrColl: any = [];
  protected m_strShapeID: string | null = null;
  protected m_xmlRunEle: Element | null = null;
  protected m_nPage = 0;
  protected m_ParaStyle = "";

  ParseTextAttributes = async (
    XmlOffice12Node: Element,
    xmlDoc: XMLDocument,
    CMLNode: Element,
    CurrentPackageReader: PackageReader
  ) => {
    this.m_xmlParaEle = XmlOffice12Node;
    this.m_myXmlDoc = xmlDoc;
    this.m_CMLProps = CMLNode;
    this.m_myPackageReader = CurrentPackageReader;

    this.SetThemeMinorFont();

    const strDocUriString = await this.m_myPackageReader.ReturnBaseXML(
      null,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    );

    this.m_docUri = CommonFunctions.PrependStringToURIPath(
      strDocUriString,
      "/"
    );

    //XmlElement fontattribs = ParseTextAttributes(m_xmlParaEle, CMLNode);
    const fontattribs = this.ParseTextAttributes1(this.m_xmlParaEle, CMLNode); // Assuming ParseTextAttributes function exists in TypeScript and handles XML nodes

    const xAuto = xmlDoc.getElementsByTagName("autoclr")[0]; // Simple tag selection, assuming a direct child path without namespace
    const autoclr = xAuto ? xAuto.textContent : "";
    this.PostProcessFontAttribs(fontattribs, autoclr || "");
  };

  GenerateParaText = (XmlOffice12Node: Element, paraProps: Element) => {
    let strParaText: string = "",
      nFormulaCount: number = 0,
      sFormula: string = "",
      sFldFlag: string = "",
      sBuffer: string = "",
      sShapeIDFormulaNode: Element | null = null,
      xinfo: Element | null = null,
      bFormulaDone: boolean = true;

    this.m_strShapeID = null;

    let objNode = XmlOffice12Node.firstChild as Element;

    while (objNode) {
      switch (objNode.nodeName) {
        case "w:r":
          const xFldChar = objNode.getElementsByTagName("w:fldChar")[0];
          if (xFldChar) {
            sFldFlag = xFldChar.getAttribute("w:fldCharType") || "";
            if (sFldFlag == "begin" && bFormulaDone == true) {
              nFormulaCount++;
              bFormulaDone = false;
            }
          }

          const xObject = objNode.getElementsByTagName("w:object")[0];
          if (xObject) {
            const xOLEObject = objNode.getElementsByTagName("o:OLEObject")[0];
            if (xOLEObject == null) {
              const xShape = objNode.getElementsByTagName("v:shape")[0];
              if (xShape) {
                const sShapeID = xShape.getAttribute("id");
                if (sShapeID && sShapeID.length > 0) {
                  this.m_strShapeID = sShapeID;
                }
              }
            }
          }
          switch (sFldFlag) {
            case "begin":
              const xInstrText = objNode.getElementsByTagName("w:instrText")[0];
              if (xInstrText) {
                sFormula += xInstrText.textContent;
              }
              // Parsing ffData node for additional information
              const xffData =
                objNode.getElementsByTagName("w:fldChar/w:ffData")[0]; // need to test this.
              if (xffData) {
                if (!this.m_myXmlDoc) {
                  throw new Error("Document not parsed!!!!");
                }

                xinfo = this.m_myXmlDoc.createElement("info");

                const xName = objNode.getElementsByTagName("w:name")[0];
                if (xName) {
                  const sName = xName.getAttribute("w:val");
                  Utils.CreateNode(
                    this.m_myXmlDoc,
                    xinfo,
                    "bookmarkname",
                    sName || ""
                  );
                }

                const xEnabled = objNode.getElementsByTagName("w:enabled")[0];
                if (xEnabled) {
                  let sEnabled = xEnabled.getAttribute("w:val");
                  if (sEnabled == null || sEnabled == "") {
                    sEnabled = "1";
                  }
                  Utils.CreateNode(this.m_myXmlDoc, xinfo, "enabled", sEnabled);
                }

                const xCalcOnExit =
                  objNode.getElementsByTagName("w:calcOnExit")[0];
                if (xCalcOnExit) {
                  let sCalcOnExit = xCalcOnExit.getAttribute("w:val");
                  if (sCalcOnExit == null || sCalcOnExit == "") {
                    sCalcOnExit = "1";
                  }
                  Utils.CreateNode(
                    this.m_myXmlDoc,
                    xinfo,
                    "calculateonexit",
                    sCalcOnExit
                  );
                }

                const xStatustext =
                  objNode.getElementsByTagName("w:statusText")[0];
                if (xStatustext) {
                  let xStatustextnode: Element = Utils.CreateNode(
                    this.m_myXmlDoc,
                    xinfo,
                    "statustext",
                    ""
                  );
                  const sStatustexttype = xStatustext.getAttribute("w:type");
                  Utils.CreateNode(
                    this.m_myXmlDoc,
                    xStatustextnode,
                    "type",
                    sStatustexttype || ""
                  );
                  const sStatustextvalue = xStatustext.getAttribute("w:val");
                  Utils.CreateNode(
                    this.m_myXmlDoc,
                    xStatustextnode,
                    "value",
                    sStatustextvalue || ""
                  );
                }

                // Parse checkbox properties
                const xCheckbox: Element =
                  objNode.getElementsByTagName("w:checkBox")[0];
                if (xCheckbox) {
                  const xCheckboxnode = Utils.CreateNode(
                    this.m_myXmlDoc,
                    xinfo,
                    "checkbox",
                    ""
                  );

                  let sCheckboxsize = "";
                  const xCheckboxsizeauto =
                    xCheckbox.getElementsByTagName("w:checkBox")[0];
                  if (xCheckboxsizeauto) {
                    sCheckboxsize = "auto";
                  } else {
                    let xCheckboxsize =
                      xCheckbox.getElementsByTagName("w:size")[0];
                    if (xCheckboxsize) {
                      sCheckboxsize = xCheckboxsize.getAttribute("w:val") || "";
                    }
                  }
                  Utils.CreateNode(
                    this.m_myXmlDoc,
                    xCheckboxnode,
                    "checkboxsize",
                    sCheckboxsize
                  );

                  let sCheckboxdefault = "";
                  const xCheckboxdefault =
                    xCheckbox.getElementsByTagName("w:default")[0];
                  if (xCheckboxdefault != null) {
                    sCheckboxdefault =
                      xCheckboxdefault.getAttribute("w:val") || "";
                  }
                  Utils.CreateNode(
                    this.m_myXmlDoc,
                    xCheckboxnode,
                    "checkboxdefault",
                    sCheckboxdefault
                  );
                } else {
                  // Check for textinput node
                  const xTextinput =
                    xffData.getElementsByTagName("w:textInput")[0];
                  if (xTextinput) {
                    const xTextinputnode = Utils.CreateNode(
                      this.m_myXmlDoc,
                      xinfo,
                      "textbox",
                      ""
                    );
                    const xType = xTextinput.getElementsByTagName("w:type")[0];
                    if (xType) {
                      const sTextboxtype = xType.getAttribute("w:val") || "";
                      Utils.CreateNode(
                        this.m_myXmlDoc,
                        xTextinputnode,
                        "type",
                        sTextboxtype
                      );
                    } else {
                      Utils.CreateNode(
                        this.m_myXmlDoc,
                        xTextinputnode,
                        "type",
                        "text"
                      );
                    }
                    const xTextboxdefault =
                      xTextinput.getElementsByTagName("w:default")[0];
                    if (xTextboxdefault) {
                      const sTextboxdefault =
                        xTextboxdefault.getAttribute("w:val") || "";
                      Utils.CreateNode(
                        this.m_myXmlDoc,
                        xTextinputnode,
                        "default",
                        sTextboxdefault
                      );
                    }
                    const xMaxlength =
                      xTextinput.getElementsByTagName("w:maxLength")[0];
                    if (xMaxlength) {
                      const sTextboxmaxlength =
                        xMaxlength.getAttribute("w:val") || "";
                      Utils.CreateNode(
                        this.m_myXmlDoc,
                        xTextinputnode,
                        "maxlength",
                        sTextboxmaxlength
                      );
                    }

                    const xFormat =
                      xTextinput.getElementsByTagName("w:format")[0];
                    if (xFormat) {
                      const sTextboxformat =
                        xFormat.getAttribute("w:val") || "";
                      Utils.CreateNode(
                        this.m_myXmlDoc,
                        xTextinputnode,
                        "format",
                        sTextboxformat
                      );
                    }
                  } else {
                    // Check for Dropdown node
                    const xDropdown =
                      xffData.getElementsByTagName("w:ddList")[0];
                    if (xDropdown) {
                      const xDropdownnode = Utils.CreateNode(
                        this.m_myXmlDoc,
                        xinfo,
                        "dropdown",
                        ""
                      );
                      const xListentry =
                        xDropdown.getElementsByTagName("w:listEntry");
                      for (let elementIndx in xListentry) {
                        const entry = Utils.CreateNode(
                          this.m_myXmlDoc,
                          xDropdownnode,
                          "entry",
                          ""
                        );
                        Utils.CreateNode(
                          this.m_myXmlDoc,
                          entry,
                          "value",
                          xListentry[elementIndx].getAttribute("w:val") || ""
                        );
                      }
                    }
                  }
                }
              }
              break;
            case "separate":
              if (bFormulaDone == false && nFormulaCount > 0) {
                const xFomulas = WordUtils.returnNode(
                  this.m_myXmlDoc as XMLDocument,
                  paraProps,
                  "formulas"
                );
                const nPos = strParaText.length;
                const xFormula = Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xFomulas,
                  "formula",
                  ""
                );
                Utils.CreateAttribute(
                  this.m_myXmlDoc as XMLDocument,
                  xFormula,
                  "id",
                  nFormulaCount.toString()
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xFormula,
                  "pos",
                  nPos.toString()
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xFormula,
                  "val",
                  sFormula
                );
                if (xinfo) {
                  xFormula.appendChild(xinfo);
                }
                sShapeIDFormulaNode = xFormula;

                strParaText += "@formula@";
              }
              sFormula = "";
              sFldFlag = "processing";
              bFormulaDone = true;
              sBuffer = "";
              xinfo = null;
              break;
            case "processing":
              if (objNode.nextSibling == null) {
                strParaText += sBuffer;
                sFldFlag = "";
              }

              const sText = this.ExtractParaTextAndSpecialcharsFromRNode(
                objNode,
                strParaText
              );

              sBuffer += sText;
              let xFrm = xpath.select1(
                `formulas/formula[@id='${nFormulaCount.toString()}']`,
                paraProps
              );
              // console.log("paraProps: ", paraProps.textContent);
              // const xFrm = paraProps.getElementsByTagName(
              //   "formulas/formula[@id='" + nFormulaCount.toString() + "']"
              // )[0];
              if (xFrm)
                WordUtils.returnNode(
                  this.m_myXmlDoc as XMLDocument,
                  xFrm as Element,
                  "infotxt"
                ).textContent = sBuffer;
              break;
            case "end":
              if (bFormulaDone == false && nFormulaCount > 0) {
                const xFomulas = WordUtils.returnNode(
                  this.m_myXmlDoc as XMLDocument,
                  paraProps,
                  "formulas"
                );
                const nPos = strParaText.length;
                const xFormula = Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xFomulas,
                  "formula",
                  ""
                );
                Utils.CreateAttribute(
                  this.m_myXmlDoc as XMLDocument,
                  xFormula,
                  "id",
                  nFormulaCount.toString()
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xFormula,
                  "pos",
                  nPos.toString()
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xFormula,
                  "val",
                  sFormula
                );
                if (this.m_strShapeID != null) {
                  Utils.CreateNode(
                    this.m_myXmlDoc as XMLDocument,
                    xFomulas,
                    "shapeid",
                    this.m_strShapeID
                  );
                }
                if (xinfo != null) {
                  xFormula.appendChild(xinfo);
                }

                strParaText += "@formula@";
              } else {
                if (sShapeIDFormulaNode != null) {
                  if (this.m_strShapeID != null) {
                    Utils.CreateNode(
                      this.m_myXmlDoc as XMLDocument,
                      sShapeIDFormulaNode,
                      "shapeid",
                      this.m_strShapeID
                    );
                  }
                }
              }
              sFormula = "";
              sFldFlag = "";
              xinfo = null;
              bFormulaDone = true;
              this.m_strShapeID = null;
              sShapeIDFormulaNode = null;
              break;
          }
          if (sFldFlag.length <= 0) {
            const sText = this.ExtractParaTextAndSpecialcharsFromRNode(
              objNode as Element,
              strParaText
            );
            strParaText += sText;
          }
          break;

        default:
          break;
      }
      objNode = objNode.nextSibling as Element;
    }
    Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      paraProps,
      "txt",
      strParaText
    );
  };

  PostProcessFontAttribs(fontattribs: Element, autoclr: string): void {
    let xList = fontattribs.getElementsByTagName("sz");
    for (let i = 0; i < xList.length; i++) {
      xList[i].textContent = WordUtils.ConvertValues(
        2,
        xList[i].textContent || "",
        true
      );
    }

    xList = fontattribs.getElementsByTagName("spac");
    for (let i = 0; i < xList.length; i++) {
      xList[i].textContent = WordUtils.ConvertValues(
        20,
        xList[i].textContent || "",
        true
      );
    }

    xList = fontattribs.getElementsByTagName("kern");
    for (let i = 0; i < xList.length; i++) {
      xList[i].textContent = WordUtils.ConvertValues(
        2,
        xList[i].textContent || "",
        true
      );
    }

    xList = fontattribs.getElementsByTagName("pos");
    for (let i = 0; i < xList.length; i++) {
      xList[i].textContent = WordUtils.ConvertValues(
        2,
        xList[i].textContent || "",
        true
      );
    }

    xList = fontattribs.getElementsByTagName("u");
    for (let i = 0; i < xList.length; i++) {
      let clr = xList[i].getElementsByTagName("clr")[0];
      if (clr && clr.textContent === "Automatic") {
        clr.textContent = autoclr;
      }
    }

    xList = fontattribs.getElementsByTagName("clr");
    for (let i = 0; i < xList.length; i++) {
      let hex = xList[i].getElementsByTagName("hex")[0];
      if (hex && hex.textContent === "Automatic") {
        hex.textContent = autoclr;
      }
    }

    xList = fontattribs.getElementsByTagName("shading");
    for (let i = 0; i < xList.length; i++) {
      let patternclr = xList[i].getElementsByTagName("patternclr")[0];
      if (patternclr && patternclr.textContent === "Automatic") {
        patternclr.textContent = autoclr;
      }
      let patternstyle = xList[i].getElementsByTagName("patternstyle")[0];
      if (patternstyle) {
        patternstyle.textContent = WordUtils.GetPatternStyle(
          patternstyle.textContent || ""
        );
      }
    }

    let hclr = "";
    xList = fontattribs.getElementsByTagName("hclr");
    for (let i = 0; i < xList.length; i++) {
      hclr = xList[i].textContent || "";
      switch (hclr) {
        case "green":
          xList[i].textContent = "Bright Green";
          break;
        // Add other cases similarly
        default:
          break;
      }
    }

    // Pagination
    xList = fontattribs.getElementsByTagName("pagination");
    for (let i = 0; i < xList.length; i++) {
      if (xList[i].textContent !== "0") {
        xList[i].textContent = "1";
      }
    }
  }

  WriteTextAttributeForPara = (
    strAttrName: string,
    paraTextAttr: Element
  ): void => {
    if (strAttrName == "valign") {
      // debugger;
    }
    let runCount: number = this.m_paraAttrColl.length;

    let strTextToWriteOut: string = "";
    let strAttributeBeingComparedValue: string = "";

    let bWriteAtEnd: boolean = false;
    let bFullyMatches: boolean = true;
    let nCurrPos: number = 0;

    let colorAttribsThis: ColorAttribs = new ColorAttribs();
    let gradStopArray: GradientStopAttributes[] | null = null;
    let shadeAttribsThis: ShadeAttribs = new ShadeAttribs();
    let underlineAttribsThis: UnderlineAttribs = new UnderlineAttribs();
    let textOutlineAttribsThis: TextOutlineAttribs = new TextOutlineAttribs();
    let shadowAttribsThis: ShadowAttribs = new ShadowAttribs();

    let ppThis: paraFontAttributes;
    let ppNext: paraFontAttributes;

    let strType: string = "solid";

    for (let i = 0; i < runCount; i++) {
      ppThis = this.m_paraAttrColl[i] as paraFontAttributes;

      strType = ppThis.strType ?? "solid";

      if (strTextToWriteOut.length === 0) {
        let strBuilderThisPara: string = strTextToWriteOut;
        strBuilderThisPara += ppThis.strText;
        strTextToWriteOut = strBuilderThisPara;
      }

      if (strAttrName === "u") {
        underlineAttribsThis.strIsUnderline = ppThis.strIsUnderline;
        underlineAttribsThis.strUnderlineClr = ppThis.strUnderlineClr;
        underlineAttribsThis.strUnderlineTClr = ppThis.strUnderlineClrThemeClr;
        underlineAttribsThis.strUnderlineTShade =
          ppThis.strUnderlineClrThemeShade;
        underlineAttribsThis.strUnderlineTTint =
          ppThis.strUnderlineClrThemeTint;
        underlineAttribsThis.strUnderlineType = ppThis.strUnderlineType;
      } else if (strAttrName === "clr") {
        if (strType === "gradient") {
          gradStopArray = ppThis.gradientStops;
        } else {
          colorAttribsThis.strClr = ppThis.strColor;
          colorAttribsThis.strTClr = ppThis.strThemeColor;
          colorAttribsThis.strTShade = ppThis.strThemeShade;
          colorAttribsThis.strTTint = ppThis.strThemeTint;
          colorAttribsThis.strType = ppThis.strType;
        }
      } else if (strAttrName === "shading") {
        shadeAttribsThis.strShdPattStyle = ppThis.strShdPattStyle;
        shadeAttribsThis.strShdPattCol = ppThis.strShdPattCol;
        shadeAttribsThis.strShdPattThemeCol = ppThis.strShdPattThemeClr;
        shadeAttribsThis.strShdPattThemeShade = ppThis.strShdPattThemeShade;
        shadeAttribsThis.strShdPattThemeTint = ppThis.strShdPattThemeTint;
        shadeAttribsThis.strShdClr = ppThis.strShdColor;
        shadeAttribsThis.strShdTClr = ppThis.strShdThemeColor;
        shadeAttribsThis.strShdTShade = ppThis.strShdThemeShade;
        shadeAttribsThis.strShdTTint = ppThis.strShdThemeTint;
      } else if (strAttrName === "txtoutline") {
        textOutlineAttribsThis.strThemColor = ppThis.strTextOutlineThemColor;
        textOutlineAttribsThis.strThemeShade = ppThis.strTextOutlineThemeShade;
        textOutlineAttribsThis.strSatMode = ppThis.strTextOutlineSatMode;
        textOutlineAttribsThis.strLineType = ppThis.strTextOutlineLineType;
        textOutlineAttribsThis.strWidth = ppThis.strTextOutlineWidth;
      } else if (strAttrName === "shadow") {
        shadowAttribsThis.strShdowblur = ppThis.strShdowblur;
        shadowAttribsThis.strShdowdist = ppThis.strShdowdist;
        shadowAttribsThis.strShdowdir = ppThis.strShdowdir;
        shadowAttribsThis.strShdowsx = ppThis.strShdowsx;
        shadowAttribsThis.strShdowsy = ppThis.strShdowsy;
        shadowAttribsThis.strShdowSchemeclr = ppThis.strShdowSchemeclr;
        shadowAttribsThis.strShdowLummod = ppThis.strShdowLummod;
        shadowAttribsThis.strShdowLumOff = ppThis.strShdowLumOff;
      }
      strAttributeBeingComparedValue = this.GetValueOfFontParameterFromStruct(
        ppThis,
        strAttrName
      );

      if (i + 1 == runCount) {
        // We dont have a next element.. this is last one.. so we will write at end..
        bWriteAtEnd = true;
      } else {
        // Get the next element to compare with
        ppNext = this.m_paraAttrColl[i + 1];

        // Earlier we were not treating the underline attribute separately.
        // Now the underline attribute is a mix of the style and color attributes
        // So use a flag "processFurther" to appropriately continue earlier processing.
        let bProcessFurther: boolean = false;

        if (strAttrName === "u") {
          if (
            ppThis.strIsUnderline === ppNext.strIsUnderline &&
            ppThis.strUnderlineClr === ppNext.strUnderlineClr &&
            ppThis.strUnderlineClrThemeClr === ppNext.strUnderlineClrThemeClr &&
            ppThis.strUnderlineClrThemeShade ===
              ppNext.strUnderlineClrThemeShade &&
            ppThis.strUnderlineClrThemeTint ===
              ppNext.strUnderlineClrThemeTint &&
            ppThis.strUnderlineType === ppNext.strUnderlineType
          ) {
            bProcessFurther = true;
          }
        } else if (strAttrName === "clr") {
          const strType: string = "gradiant"; // Example type
          if (
            strType === "gradiant" &&
            this.CompareGradientArrays(
              ppThis.gradientStops,
              ppNext.gradientStops
            )
          ) {
            bProcessFurther = true;
          } else if (
            ppThis.strColor === ppNext.strColor &&
            ppThis.strThemeColor === ppNext.strThemeColor &&
            ppThis.strThemeShade === ppNext.strThemeShade &&
            ppThis.strThemeTint === ppNext.strThemeTint
          ) {
            bProcessFurther = true;
          }
        } else if (strAttrName === "shading") {
          if (
            ppThis.strShdPattStyle === ppNext.strShdPattStyle &&
            ppThis.strShdPattCol === ppNext.strShdPattCol &&
            ppThis.strShdPattThemeClr === ppNext.strShdPattThemeClr &&
            ppThis.strShdPattThemeShade === ppNext.strShdPattThemeShade &&
            ppThis.strShdPattThemeTint === ppNext.strShdPattThemeTint &&
            ppThis.strShdColor === ppNext.strShdColor &&
            ppThis.strShdThemeColor === ppNext.strShdThemeColor &&
            ppThis.strShdThemeColor === ppNext.strShdThemeShade &&
            ppThis.strShdThemeTint === ppNext.strShdThemeTint
          ) {
            bProcessFurther = true;
          }
        } else if (strAttrName === "txtoutline") {
          if (
            ppThis.strTextOutlineThemColor === ppNext.strTextOutlineThemColor &&
            ppThis.strTextOutlineThemeShade ===
              ppNext.strTextOutlineThemeShade &&
            ppThis.strTextOutlineSatMode === ppNext.strTextOutlineSatMode &&
            ppThis.strTextOutlineLineType === ppNext.strTextOutlineLineType &&
            ppThis.strTextOutlineWidth === ppNext.strTextOutlineWidth
          ) {
            bProcessFurther = true;
          }
        } else if (strAttrName === "shadow") {
          if (
            ppThis.strShdowblur === ppNext.strShdowblur &&
            ppThis.strShdowdist === ppNext.strShdowdist &&
            ppThis.strShdowdir === ppNext.strShdowdir &&
            ppThis.strShdowsx === ppNext.strShdowsx &&
            ppThis.strShdowsy === ppNext.strShdowsy &&
            ppThis.strShdowSchemeclr === ppNext.strShdowSchemeclr &&
            ppThis.strShdowLummod === ppNext.strShdowLummod &&
            ppThis.strShdowLumOff === ppNext.strShdowLumOff
          ) {
            bProcessFurther = true;
          }
        } else if (
          strAttributeBeingComparedValue ==
          this.GetValueOfFontParameterFromStruct(ppNext, strAttrName)
        ) {
          bProcessFurther = true;
        }

        if (bProcessFurther) {
          bWriteAtEnd = true;

          let strBuilderNextPara: string = "";

          strBuilderNextPara += strTextToWriteOut;
          strBuilderNextPara += ppNext.strText;

          strTextToWriteOut = strBuilderNextPara;
        } else {
          // let xmlEle = xpath.select1(
          //   `*[local-name()='${strAttrName}'][@pos='${nCurrPos.toString()}']`,
          //   paraTextAttr
          // ) as Element | null;
          bWriteAtEnd = false;
          bFullyMatches = false;
          let xmlEle = xpath.select1(
            strAttrName + "[@pos='" + nCurrPos.toString() + "']",
            paraTextAttr
          ) as Element | null;

          // Temporary fix for break
          if (xmlEle) {
            if (
              (strAttrName === "br" && xmlEle.textContent === "none") ||
              strAttrName !== "br"
            ) {
              while (xmlEle.firstChild) {
                xmlEle.removeChild(xmlEle.firstChild);
              }
            } else {
              xmlEle = Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                paraTextAttr,
                strAttrName,
                ""
              );
            }
          } // Dont generate br node in case its value is none
          else if (
            strAttrName == "br" &&
            strAttributeBeingComparedValue == "none"
          ) {
            xmlEle = null;
          } else {
            xmlEle = Utils.CreateNode(
              this.m_myXmlDoc as XMLDocument,
              paraTextAttr,
              strAttrName,
              ""
            );
          }
          if (nCurrPos == 0) {
            if (strAttrName === "clr") {
              if (ppThis.strType === "gradiant") {
                this.WriteAttributesForGradientColorNode(
                  xmlEle as Element,
                  strTextToWriteOut,
                  "1",
                  false,
                  gradStopArray as GradientStopAttributes[]
                );
              } else {
                this.WriteAttributesForTextColorNode(
                  xmlEle as Element,
                  strTextToWriteOut,
                  "1",
                  false,
                  colorAttribsThis
                );
              }
            } else if (strAttrName === "shading") {
              this.WriteAttributesForTextShadeNode(
                xmlEle as Element,
                strTextToWriteOut,
                "1",
                false,
                shadeAttribsThis
              );
            } else if (strAttrName === "u") {
              this.WriteAttributesForTextUnderlineNode(
                xmlEle as Element,
                strTextToWriteOut,
                "1",
                false,
                underlineAttribsThis
              );
            } else if (strAttrName === "txtoutline") {
              this.WriteAttributesForTextOutlineNode(
                xmlEle as Element,
                strTextToWriteOut,
                "1",
                false,
                textOutlineAttribsThis
              );
            } else if (strAttrName === "shadow") {
              this.WriteAttributesForShadowNode(
                xmlEle as Element,
                strTextToWriteOut,
                "1",
                false,
                shadowAttribsThis
              );
            } else {
              if (xmlEle) {
                xmlEle.textContent = strAttributeBeingComparedValue;
                this.GetCommonAttributes(xmlEle, strTextToWriteOut, "1", false);
              }
            }
          } else {
            if (strAttrName === "clr") {
              if (ppThis.strType === "gradiant") {
                this.WriteAttributesForGradientColorNode(
                  xmlEle as Element,
                  strTextToWriteOut,
                  nCurrPos.toString(),
                  false,
                  gradStopArray || []
                );
              } else {
                this.WriteAttributesForTextColorNode(
                  xmlEle as Element,
                  strTextToWriteOut,
                  nCurrPos.toString(),
                  false,
                  colorAttribsThis
                );
              }
            } else if (strAttrName === "shading") {
              this.WriteAttributesForTextShadeNode(
                xmlEle as Element,
                strTextToWriteOut,
                nCurrPos.toString(),
                false,
                shadeAttribsThis
              );
            } else if (strAttrName === "u") {
              this.WriteAttributesForTextUnderlineNode(
                xmlEle as Element,
                strTextToWriteOut,
                nCurrPos.toString(),
                false,
                underlineAttribsThis
              );
            } else if (strAttrName === "txtoutline") {
              this.WriteAttributesForTextOutlineNode(
                xmlEle as Element,
                strTextToWriteOut,
                nCurrPos.toString(),
                false,
                textOutlineAttribsThis
              );
            } else if (strAttrName === "shadow") {
              this.WriteAttributesForShadowNode(
                xmlEle as Element,
                strTextToWriteOut,
                nCurrPos.toString(),
                false,
                shadowAttribsThis
              );
            } else {
              if (xmlEle != null) {
                xmlEle.textContent = strAttributeBeingComparedValue;
                this.GetCommonAttributes(
                  xmlEle,
                  strTextToWriteOut,
                  nCurrPos.toString(),
                  false
                );
              }
            }
          }

          //there can arise cases in which strTextToWriteOut.length is 0. To correctly adjust nPos, increment for this case also.
          if (nCurrPos == 0 || strTextToWriteOut.length == 0) {
            nCurrPos += strTextToWriteOut.length + 1;
          } else {
            nCurrPos += strTextToWriteOut.length;
          }

          strTextToWriteOut = "";
          strAttributeBeingComparedValue = "";
        }
      }
    }

    // Either we had only a single element OR the last consecutive ones matched...
    if (bWriteAtEnd) {
      let xmlEle: Element | null = null;

      // Construct XPath query to select node based on attribute conditions
      const strXPath = `${strAttrName}[@pos='${nCurrPos}']`;

      // Select the node using XPath
      xmlEle = xpath.select1(strXPath, paraTextAttr) as Element;

      // Check if the node exists
      if (xmlEle) {
        // Special handling for "br" node
        if (strAttrName !== "br") {
          xmlEle.childNodes.forEach((child) =>
            (xmlEle as Element).removeChild(child)
          );
        } else if (
          strAttrName === "br" &&
          strAttributeBeingComparedValue === "none"
        ) {
          xmlEle = null;
        } else {
          xmlEle = Utils.CreateNode(
            this.m_myXmlDoc as XMLDocument,
            paraTextAttr,
            strAttrName,
            ""
          );
        }
      } else {
        if (strAttrName === "br") {
          // If 1 or more "br" nodes are present and attribute value is "none", do not generate new "br" nodes
          const brNodes = xpath.select(strAttrName, paraTextAttr) as Node[];
          if (
            brNodes.length >= 1 &&
            strAttributeBeingComparedValue === "none"
          ) {
            xmlEle = null;
          } else {
            // In case no "br" is present, generate "br" node with "none" value
            xmlEle = Utils.CreateNode(
              this.m_myXmlDoc as XMLDocument,
              paraTextAttr,
              strAttrName,
              ""
            );
          }
        } else {
          xmlEle = Utils.CreateNode(
            this.m_myXmlDoc as XMLDocument,
            paraTextAttr,
            strAttrName,
            ""
          );
        }
      }

      if (bFullyMatches) {
        if (strAttrName === "clr") {
          if (strType === "gradiant") {
            this.WriteAttributesForGradientColorNode(
              xmlEle as Element,
              "",
              "1",
              true,
              gradStopArray || []
            );
          } else {
            this.WriteAttributesForTextColorNode(
              xmlEle as Element,
              "",
              "1",
              true,
              colorAttribsThis
            );
          }
        } else if (strAttrName == "shading") {
          this.WriteAttributesForTextShadeNode(
            xmlEle as Element,
            "",
            "1",
            true,
            shadeAttribsThis
          );
        } else if (strAttrName === "u") {
          this.WriteAttributesForTextUnderlineNode(
            xmlEle as Element,
            "",
            "1",
            true,
            underlineAttribsThis
          );
        } else if (strAttrName === "txtoutline") {
          this.WriteAttributesForTextOutlineNode(
            xmlEle as Element,
            "",
            "1",
            true,
            textOutlineAttribsThis
          );
        } else if (strAttrName === "shadow") {
          this.WriteAttributesForShadowNode(
            xmlEle as Element,
            "",
            "1",
            true,
            shadowAttribsThis
          );
        } else {
          if (xmlEle != null) {
            xmlEle.textContent = strAttributeBeingComparedValue;
            this.GetCommonAttributes(xmlEle, "", "1", true);
          }
        }
      } else {
        if (strAttrName === "clr") {
          if (strType === "gradiant") {
            this.WriteAttributesForGradientColorNode(
              xmlEle as Element,
              strTextToWriteOut,
              nCurrPos.toString(),
              false,
              gradStopArray || []
            );
          } else {
            this.WriteAttributesForTextColorNode(
              xmlEle as Element,
              strTextToWriteOut,
              nCurrPos.toString(),
              false,
              colorAttribsThis
            );
          }
        } else if (strAttrName === "shading") {
          this.WriteAttributesForTextShadeNode(
            xmlEle as Element,
            strTextToWriteOut,
            nCurrPos.toString(),
            false,
            shadeAttribsThis
          );
        } else if (strAttrName === "u") {
          this.WriteAttributesForTextUnderlineNode(
            xmlEle as Element,
            strTextToWriteOut,
            nCurrPos.toString(),
            false,
            underlineAttribsThis
          );
        } else if (strAttrName === "txtoutline") {
          this.WriteAttributesForTextOutlineNode(
            xmlEle as Element,
            strTextToWriteOut,
            nCurrPos.toString(),
            false,
            textOutlineAttribsThis
          );
        } else if (strAttrName === "shadow") {
          this.WriteAttributesForShadowNode(
            xmlEle as Element,
            strTextToWriteOut,
            nCurrPos.toString(),
            false,
            shadowAttribsThis
          );
        } else {
          if (xmlEle != null) {
            xmlEle.textContent = strAttributeBeingComparedValue;
            this.GetCommonAttributes(
              xmlEle,
              strTextToWriteOut,
              nCurrPos.toString(),
              false
            );
          }
        }
      }

      const xList = paraTextAttr.getElementsByTagName(strAttrName);

      if (xList.length === 1) {
        if (xList[0].tagName !== "br") {
          xList[0].setAttribute("sameasparatxt", "1");
          xList[0].setAttribute("txt", "");
        }
      }
    }
  };

  protected ReturnTStyleNode(): Element | null {
    let xmlTStyleNode: Element | null = null;

    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });

    const xmlStyleEle = select(
      "../../../w:tblPr",
      this.m_xmlParaEle as Node,
      true
    ) as Element;

    if (xmlStyleEle) {
      const strStyleName = WordUtils.extractStringFromVal(
        xmlStyleEle,
        "w:tblStyle",
        select
      );
      const strXPath = `/w:styles/w:style[@w:type='table' and @w:styleId='${strStyleName}']`;

      if (!this.m_myPackageReader) {
        throw new Error("Package reader not initialized!!");
      }

      const doc = this.m_myPackageReader.partFileMap.get(
        "m_StylesDoc"
      ) as XMLDocument;
      xmlTStyleNode = select(strXPath, doc, true) as Element;

      if (!xmlTStyleNode) {
        return null;
      }
    }
    return xmlTStyleNode;
  }

  protected GetTableRelatedNodes(XmlOffice12Node: Element, CMLPProps: Element) {
    // Initialize namespaces if necessary
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });

    // First check if the rPr node is present in the table style
    let xmlTStyleNode = this.ReturnTStyleNode();
    if (xmlTStyleNode) {
      let xmlrPrEle = select("w:rPr", xmlTStyleNode, true) as Element;
      if (xmlrPrEle) {
        let paramarkAttr = new ParaMarkAttributes2013();
        paramarkAttr.parseParaMarkAttributes(
          xmlrPrEle,
          this.m_myXmlDoc as XMLDocument,
          CMLPProps,
          this.m_myPackageReader,
          "paramark-fontattribs"
        );
      }
    }

    // First check if the CNF node is present in the pPr node
    let xmlPProps = select("w:pPr", XmlOffice12Node, true) as Element;
    let strCnfStyleFrompPr = WordUtils.extractStringFromVal(
      xmlPProps,
      "w:cnfStyle",
      select
    );

    if (strCnfStyleFrompPr !== "") {
      let xmlTblStyleNode = this.ExtractTableStyleNode(strCnfStyleFrompPr);
      if (xmlTblStyleNode) {
        let xmlrPrEle = select("w:rPr", xmlTblStyleNode, true) as Element;
        if (xmlrPrEle) {
          let paramarkAttr = new ParaMarkAttributes2013();
          paramarkAttr.parseParaMarkAttributes(
            xmlrPrEle,
            this.m_myXmlDoc as XMLDocument,
            CMLPProps,
            this.m_myPackageReader,
            "paramark-fontattribs"
          );
        }
      }
    }

    // Now check if the Cnf is present in rPr node
    let xCnfStyle = select(
      "parent::node()/w:tcPr/w:cnfStyle",
      XmlOffice12Node,
      true
    ) as Element;

    if (!xCnfStyle) {
      // Now check if the Cnf is present in trPr node
      xCnfStyle = select(
        "parent::node()/parent::node()/w:trPr/w:cnfStyle",
        XmlOffice12Node,
        true
      ) as Element;
      if (!xCnfStyle) {
        return;
      }
    }

    let attrColl = xCnfStyle.attributes;
    if (attrColl) {
      let strCnfStyle = attrColl.getNamedItem("w:val")
        ? attrColl.getNamedItem("w:val")?.value
        : "";
      if (strCnfStyle !== "") {
        let xmlTblStyleNode = this.ExtractTableStyleNode(strCnfStyle || "");
        if (xmlTblStyleNode) {
          let xmlrPrEle = select("w:rPr", xmlTblStyleNode, true) as Element;
          if (xmlrPrEle) {
            let paramarkAttr = new ParaMarkAttributes2013();
            paramarkAttr.parseParaMarkAttributes(
              xmlrPrEle,
              this.m_myXmlDoc as XMLDocument,
              CMLPProps,
              this.m_myPackageReader,
              "paramark-fontattribs"
            );
          }
        }
      }
    }
  }

  protected ParseTextAttributes1 = (
    xmlEle: Element,
    pProps: Element
  ): Element => {
    const paraTextAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      pProps,
      "fontattribs",
      ""
    );

    //Fisrt Check for the Font Attribs in the Table style
    if (this.m_bParaInCell) {
      this.GetTableRelatedNodes(this.m_xmlParaEle as Element, pProps);
    }
    //Then store the font attributes run wise in an array
    this.FillParaAttrCollection(xmlEle);

    this.WriteTextAttributeForPara("name", paraTextAttr);
    this.WriteTextAttributeForPara("sz", paraTextAttr);
    this.WriteTextAttributeForPara("b", paraTextAttr);
    this.WriteTextAttributeForPara("i", paraTextAttr);
    this.WriteTextAttributeForPara("outline", paraTextAttr);
    this.WriteTextAttributeForPara("emboss", paraTextAttr);
    this.WriteTextAttributeForPara("engrave", paraTextAttr);
    this.WriteTextAttributeForPara("valign", paraTextAttr);
    this.WriteTextAttributeForPara("shadow", paraTextAttr);
    this.WriteTextAttributeForPara("hidden", paraTextAttr);
    this.WriteTextAttributeForPara("smallcaps", paraTextAttr);
    this.WriteTextAttributeForPara("caps", paraTextAttr);

    this.WriteTextAttributeForPara("pagination", paraTextAttr);
    this.WriteTextAttributeForPara("strike", paraTextAttr);
    this.WriteTextAttributeForPara("dstrike", paraTextAttr);
    this.WriteTextAttributeForPara("u", paraTextAttr);
    this.WriteTextAttributeForPara("clr", paraTextAttr);
    this.WriteTextAttributeForPara("shading", paraTextAttr);
    this.WriteTextAttributeForPara("txtoutline", paraTextAttr);
    this.WriteTextAttributeForPara("hclr", paraTextAttr);
    this.WriteTextAttributeForPara("style", paraTextAttr);
    this.WriteTextAttributeForPara("spac", paraTextAttr);
    this.WriteTextAttributeForPara("wd", paraTextAttr);
    this.WriteTextAttributeForPara("kern", paraTextAttr);
    this.WriteTextAttributeForPara("pos", paraTextAttr);
    this.WriteTextAttributeForPara("footnoteid", paraTextAttr);
    this.WriteTextAttributeForPara("endnoteid", paraTextAttr);
    this.WriteTextAttributeForPara("br", paraTextAttr);
    this.WriteTextAttributeForPara("autoupdate", paraTextAttr);
    this.WriteTextAttributeForPara("styletype", paraTextAttr);
    this.WriteTextAttributeForPara("ligature", paraTextAttr);

    return paraTextAttr;
  };

  protected GetValueOfFontParameterFromStruct = (
    pp: paraFontAttributes,
    strAttrName: string
  ): string => {
    let strRetVal = "";
    switch (strAttrName) {
      case "name":
        strRetVal = pp.strFontName || this.m_strThemeMinorFont || "";
        break;
      case "ligature":
        strRetVal = pp.strLigature || "none";
        break;
      case "sz":
        strRetVal = pp.strFontSize || "20";
        break;
      case "clr":
        strRetVal = pp.strColor || "000000";
        break;
      case "shading":
        strRetVal = pp.strShdColor || "000000";
        break;
      case "b":
        strRetVal = pp.strBold || "0";
        break;
      case "autoupdate":
        strRetVal = pp.strAutoUpdate || "0";
        break;
      case "styletype":
        strRetVal = pp.strStyleType || "none";
        break;
      case "u":
        strRetVal = pp.strIsUnderline || "0";
        break;
      case "i":
        strRetVal = pp.strItalic || "0";
        break;
      case "valign":
        strRetVal = pp.strVAlign || "0";
        break;
      case "strike":
        strRetVal = pp.strStrikethrough || "0";
        break;
      case "dstrike":
        strRetVal = pp.strDoubleStrike || "0";
        break;
      case "pagination":
        strRetVal = pp.strPagination || "0";
        break;
      case "caps":
        strRetVal = pp.strCaps || "0";
        break;
      case "smallcaps":
        strRetVal = pp.strSmallCaps || "0";
        break;
      case "outline":
        strRetVal = pp.strOutline || "0";
        break;
      case "emboss":
        strRetVal = pp.strEmboss || "0";
        break;
      case "engrave":
        strRetVal = pp.strEngrave || "0";
        break;
      case "shadow":
        strRetVal = pp.strShadow || "0";
        break;
      case "hidden":
        strRetVal = pp.strHidden || "0";
        break;
      case "hclr":
        strRetVal = pp.strHighlightColor || "none";
        break;
      case "style":
        strRetVal = pp.strStyle || "none";
        break;
      case "spac":
        strRetVal = pp.strSpacing || "0";
        break;
      case "pos":
        strRetVal = pp.strPosition || "0";
        break;
      case "wd":
        strRetVal = pp.strWidth || "0";
        break;
      case "kern":
        strRetVal = pp.strKerning || "0";
        break;
      case "footnoteid":
        strRetVal = pp.strFootnoteID || "none";
        break;
      case "endnoteid":
        strRetVal = pp.strEndnoteID || "none";
        break;
      case "br":
        strRetVal = pp.strBreakType || "none";
        break;
      default:
        strRetVal = "";
        break;
    }
    return strRetVal;
  };

  protected ExtractParaTextAndSpecialcharsFromRNode = (
    XmlRNode: Element,
    sParaTxt: string
  ) => {
    let strParaText = "";
    const xSCs = WordUtils2010.returnNode(
      this.m_myXmlDoc as XMLDocument,
      this.m_CMLProps as Element,
      "specialchars"
    );
    let xSC = null;
    let objNode = XmlRNode.firstChild;

    while (objNode != null) {
      switch (objNode.nodeName) {
        case "w:t":
          strParaText += objNode.textContent;

          // Read run text
          const p = objNode.textContent;

          if (!p) break;

          for (let i = 0; i < p.length; i++) {
            switch (p[i].codePointAt(0)) {
              case 8211:
                xSC = Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSCs,
                  "specialchar",
                  ""
                );
                Utils.CreateAttribute(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "id",
                  this.nIndex.toString()
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "type",
                  "EnDash"
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "pos",
                  (sParaTxt.length + i).toString()
                );
                this.nIndex++;
                break;

              case 8230:
                xSC = Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSCs,
                  "specialchar",
                  ""
                );
                Utils.CreateAttribute(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "id",
                  this.nIndex.toString()
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "type",
                  "Ellipsis"
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "pos",
                  (sParaTxt.length + i).toString()
                );
                this.nIndex++;
                break;

              case 8212:
                xSC = Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSCs,
                  "specialchar",
                  ""
                );
                Utils.CreateAttribute(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "id",
                  this.nIndex.toString()
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "type",
                  "EmDash"
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "pos",
                  (sParaTxt.length + i).toString()
                );
                this.nIndex++;
                break;

              case 160:
                xSC = Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSCs,
                  "specialchar",
                  ""
                );
                Utils.CreateAttribute(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "id",
                  this.nIndex.toString()
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "type",
                  "NonBreakingSpace"
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "pos",
                  (sParaTxt.length + i).toString()
                );
                this.nIndex++;
                break;

              case 8195:
                xSC = Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSCs,
                  "specialchar",
                  ""
                );
                Utils.CreateAttribute(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "id",
                  this.nIndex.toString()
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "type",
                  "NonBreakingSpace"
                );
                Utils.CreateNode(
                  this.m_myXmlDoc as XMLDocument,
                  xSC,
                  "pos",
                  (sParaTxt.length + i).toString()
                );
                this.nIndex++;
                break;

              default:
                break;
            }
          }
          break;
        case "w:fldSimple":
          strParaText += "@formula@";
          break;

        case "w:tab":
          strParaText += "\\t";
          break;
        case "w:noBreakHyphen":
          strParaText += "\x2011";

          xSC = Utils.CreateNode(
            this.m_myXmlDoc as XMLDocument,
            xSCs,
            "specialchar",
            ""
          );
          Utils.CreateAttribute(
            this.m_myXmlDoc as XMLDocument,
            xSC,
            "id",
            this.nIndex.toString()
          );
          Utils.CreateNode(
            this.m_myXmlDoc as XMLDocument,
            xSC,
            "type",
            "NonBreakingHyphen"
          );
          Utils.CreateNode(
            this.m_myXmlDoc as XMLDocument,
            xSC,
            "pos",
            sParaTxt.length.toString()
          );
          this.nIndex++;
          break;

        case "w:softHyphen":
          strParaText += "\x00AD";

          xSC = Utils.CreateNode(
            this.m_myXmlDoc as XMLDocument,
            xSCs,
            "specialchar",
            ""
          );
          Utils.CreateAttribute(
            this.m_myXmlDoc as XMLDocument,
            xSC,
            "id",
            this.nIndex.toString()
          );
          Utils.CreateNode(
            this.m_myXmlDoc as XMLDocument,
            xSC,
            "type",
            "OptionalHyphen"
          );
          Utils.CreateNode(
            this.m_myXmlDoc as XMLDocument,
            xSC,
            "pos",
            sParaTxt.length.toString()
          );
          this.nIndex++;
          break;
        case "w:sym":
          if ((objNode as Element).attributes != null) {
            const xChar = (objNode as Element).getAttribute("w:char");
            if (xChar != null) {
              const strChar = xChar;
              const itemp = parseInt(strChar, 16); // Convert.ToInt32(strChar, 16);
              strParaText += String.fromCharCode(itemp); //Convert.ToChar(itemp);

              let strFont = "";

              const xFont = (objNode as Element).getAttribute("w:font"); //objNode.Attributes["w:font"];
              if (xFont != null) {
                strFont = xFont;
              }

              xSC = Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                xSCs,
                "specialchar",
                ""
              );
              Utils.CreateAttribute(
                this.m_myXmlDoc as XMLDocument,
                xSC,
                "id",
                this.nIndex.toString()
              );
              Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                xSC,
                "type",
                strFont + strChar
              );
              Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                xSC,
                "font",
                strFont
              );
              Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                xSC,
                "char",
                strChar
              );
              Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                xSC,
                "pos",
                sParaTxt.length.toString()
              );
              this.nIndex++;
            }
          }
          break;
      }
      objNode = objNode.nextSibling;
    }
    return strParaText;
  };

  protected WriteAttributesForGradientColorNode = (
    xmlElementPassed: Element,
    textToWrite: string,
    positionOfText: string,
    fullyMatchedPara: boolean,
    clrAttribs: GradientStopAttributes[]
  ): void => {
    this.GetCommonAttributes(
      xmlElementPassed,
      textToWrite,
      positionOfText,
      fullyMatchedPara
    );

    let tempAttr, stops;

    Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "type",
      "gradient"
    );
    stops = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "stops",
      ""
    );

    clrAttribs.forEach((stopAttr: GradientStopAttributes) => {
      tempAttr = Utils.CreateNode(
        this.m_myXmlDoc as XMLDocument,
        stops,
        "stop",
        ""
      );
      Utils.CreateAttribute(
        this.m_myXmlDoc as XMLDocument,
        tempAttr,
        "pos",
        stopAttr.strPos
      );

      Utils.CreateNode(
        this.m_myXmlDoc as XMLDocument,
        tempAttr,
        "clr",
        stopAttr.strTClr
      );
      Utils.CreateNode(
        this.m_myXmlDoc as XMLDocument,
        tempAttr,
        "sat",
        stopAttr.strTSat
      );
      Utils.CreateNode(
        this.m_myXmlDoc as XMLDocument,
        tempAttr,
        "tint",
        stopAttr.strTTint
      );
    });
  };

  protected GetCommonAttributes = (
    xmlElementPassed: Element,
    textToWrite: string,
    positionOfText: string,
    fullyMatchedPara: boolean
  ): void => {
    //do nothing in this case
    if (xmlElementPassed == null) {
      return;
    }

    Utils.CreateAttribute(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "txt",
      textToWrite
    );
    Utils.CreateAttribute(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "pos",
      positionOfText
    );

    if (fullyMatchedPara) {
      Utils.CreateAttribute(
        this.m_myXmlDoc as XMLDocument,
        xmlElementPassed,
        "sameasparatxt",
        "1"
      );
    } else {
      Utils.CreateAttribute(
        this.m_myXmlDoc as XMLDocument,
        xmlElementPassed,
        "sameasparatxt",
        "0"
      );
    }
  };

  protected WriteAttributesForTextColorNode = (
    xmlElementPassed: Element,
    textToWrite: string,
    positionOfText: string,
    fullyMatchedPara: boolean,
    clrAttribs: ColorAttribs
  ): void => {
    this.GetCommonAttributes(
      xmlElementPassed,
      textToWrite,
      positionOfText,
      fullyMatchedPara
    );

    let tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "filltype",
      clrAttribs.strType
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "hex",
      clrAttribs.strClr
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themeclr",
      clrAttribs.strTClr
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themeshade",
      clrAttribs.strTShade
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themetint",
      clrAttribs.strTTint
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
  };

  protected WriteAttributesForTextUnderlineNode = (
    xmlElementPassed: Element,
    textToWrite: string,
    positionOfText: string,
    fullyMatchedPara: boolean,
    underlineAttribsThis: UnderlineAttribs
  ): void => {
    this.GetCommonAttributes(
      xmlElementPassed,
      textToWrite,
      positionOfText,
      fullyMatchedPara
    );

    let tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "value",
      underlineAttribsThis.strIsUnderline
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "0";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "clr",
      underlineAttribsThis.strUnderlineClr
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "type",
      underlineAttribsThis.strUnderlineType
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themeclr",
      underlineAttribsThis.strUnderlineTClr
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themeshade",
      underlineAttribsThis.strUnderlineTShade
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themetint",
      underlineAttribsThis.strUnderlineTTint
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
  };

  protected WriteAttributesForTextShadeNode = (
    xmlElementPassed: Element,
    textToWrite: string,
    positionOfText: string,
    fullyMatchedPara: boolean,
    shdAttribs: ShadeAttribs
  ): void => {
    this.GetCommonAttributes(
      xmlElementPassed,
      textToWrite,
      positionOfText,
      fullyMatchedPara
    );

    let tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "patternstyle",
      shdAttribs.strShdPattStyle
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "Clear";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "patternclr",
      shdAttribs.strShdPattCol
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "Automatic";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "patternthemeclr",
      shdAttribs.strShdPattThemeCol
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "patternthemeshade",
      shdAttribs.strShdPattThemeShade
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "patternthemetint",
      shdAttribs.strShdPattThemeTint
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "fillclr",
      shdAttribs.strShdClr
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themeclr",
      shdAttribs.strShdTClr
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themeshade",
      shdAttribs.strShdTShade
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themetint",
      shdAttribs.strShdTTint
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
  };

  private WriteAttributesForShadowNode = (
    xmlElementPassed: Element,
    textToWrite: string,
    positionOfText: string,
    fullyMatchedPara: boolean,
    shadowThis: ShadowAttribs
  ): void => {
    this.GetCommonAttributes(
      xmlElementPassed,
      textToWrite,
      positionOfText,
      fullyMatchedPara
    );

    let tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "blur",
      shadowThis.strShdowblur
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "dist",
      shadowThis.strShdowdist
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "angle",
      shadowThis.strShdowdir
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "sizex",
      shadowThis.strShdowsx
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "sizey",
      shadowThis.strShdowsy
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "default";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "schemeclr",
      shadowThis.strShdowSchemeclr
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "default";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "lumnode",
      shadowThis.strShdowLummod
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "default";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "lumoff",
      shadowThis.strShdowLumOff
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "default";
    }
  };

  private WriteAttributesForTextOutlineNode = (
    xmlElementPassed: Element,
    textToWrite: string,
    positionOfText: string,
    fullyMatchedPara: boolean,
    textOutlineAttribsThis: TextOutlineAttribs
  ) => {
    this.GetCommonAttributes(
      xmlElementPassed,
      textToWrite,
      positionOfText,
      fullyMatchedPara
    );

    let tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themeclr",
      textOutlineAttribsThis.strThemColor
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }

    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "themeshade",
      textOutlineAttribsThis.strThemeShade
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "satmod",
      textOutlineAttribsThis.strSatMode
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "linetype",
      textOutlineAttribsThis.strLineType
    );
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "none";
    }
    tempAttr = Utils.CreateNode(
      this.m_myXmlDoc as XMLDocument,
      xmlElementPassed,
      "linewidth",
      textOutlineAttribsThis.strWidth
    );
    this.m_myXmlDoc as XMLDocument;
    if (tempAttr.textContent == "") {
      tempAttr.textContent = "default";
    }
  };

  private SetThemeMinorFont = () => {
    if (!this.m_myPackageReader) {
      throw new Error("Package is not initailized yet..");
    }

    const xmlDoc = this.m_myPackageReader.partFileMap.get("m_ThemeDoc");

    if (!xmlDoc) {
      throw new Error("Theme document not Found!!!!");
    }

    let strXPath = "/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin";
    // Initialize namespace manager
    const select = xpath.useNamespaces({
      a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    });

    let xElement: Element[] = select(
      strXPath,
      xmlDoc as XMLDocument
    ) as Element[]; // XmlElement xElement = (XmlElement)xmlDoc.SelectSingleNode(strXPath, m_nsmgrA);
    if (xElement.length) {
      this.m_strThemeMinorFont = xElement[0].getAttribute("typeface"); // +"Body";
    }

    strXPath = "/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin";
    xElement = select(strXPath, xmlDoc as XMLDocument) as Element[];
    if (xElement != null) {
      this.m_strThemeMajorFont = xElement[0].getAttribute("typeface"); // +"Heading";
    }
  };

  private CompareGradientArrays = (
    gradientStopAttributesThis: GradientStopAttributes[],
    gradientStopAttributesNext: GradientStopAttributes[]
  ): boolean => {
    if (!gradientStopAttributesThis || !gradientStopAttributesNext) {
      return false;
    }

    if (
      gradientStopAttributesThis.length !== gradientStopAttributesNext.length
    ) {
      return false;
    }

    for (let nIndex = 0; nIndex < gradientStopAttributesThis.length; nIndex++) {
      const grAttrThis = gradientStopAttributesThis[nIndex];
      const grAttrNext = gradientStopAttributesNext[nIndex];

      if (
        grAttrThis.strPos !== grAttrNext.strPos ||
        grAttrThis.strTClr !== grAttrNext.strTClr ||
        grAttrThis.strTSat !== grAttrNext.strTSat ||
        grAttrThis.strTTint !== grAttrNext.strTTint
      ) {
        return false;
      }
    }

    return true;
  };

  protected ExtractTableStyleNode = (strCnfStyle: string): Element | null => {
    const xmlTStyleNode = this.ReturnTStyleNode();
    const strRef = WordUtils.ReturnCnfIndexRef(strCnfStyle);
    const strXPath = "w:tblStylePr[@w:type='" + strRef + "']";
    //this node correspond  to the w:tblStylePr, extract the properties from this first.

    if (xmlTStyleNode != null) {
      const select = xpath.useNamespaces({
        w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      });
      return select(strXPath, xmlTStyleNode, true) as Element;
    } else {
      return null;
    }
  };

  protected FillParaAttrCollection = (xmlEle: Element): void => {
    // This line will select all the Runs directly inside w:p node and w:p/w:hyperlink node
    // This is done since previous xpath was selecting the runs inside w:p/w:sdt. Because of this pargraph fontatribs node was having nodes corresponding to the format of the SDT also
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });
    const xmlList = select(
      "w:r|w:hyperlink/w:r|w:fldSimple/w:r",
      xmlEle
    ) as Element[];
    let xmlRPrEle = null;

    const nCount = this.m_paraAttrColl.length;
    const nListCount = xmlList.length;
    for (let i = 0; i < nListCount; i++) {
      let pp: paraFontAttributes;
      this.m_xmlRunEle = xmlList[i] as Element;
      if (nCount === 1 && i === 0) {
        pp = this.m_paraAttrColl[0] as paraFontAttributes;
      } else {
        pp = new paraFontAttributes();

        pp.strBold = "0";

        pp.strColor = "Automatic";
        pp.strThemeColor = "none";
        pp.strThemeShade = "none";
        pp.strThemeTint = "none";

        pp.strTextOutlineThemColor = "none";
        pp.strTextOutlineThemeShade = "none";
        pp.strTextOutlineSatMode = "none";
        pp.strTextOutlineLineType = "none";
        pp.strTextOutlineWidth = "default";

        pp.strShdColor = "No Color";
        pp.strShdThemeColor = "none";
        pp.strShdThemeShade = "none";
        pp.strShdThemeTint = "none";
        pp.strShdPattStyle = "Clear";
        pp.strShdPattCol = "Automatic";
        pp.strShdPattThemeClr = "none";
        pp.strShdPattThemeShade = "none";
        pp.strShdPattThemeTint = "none";

        pp.strFontName = this.m_strThemeMinorFont || "";
        pp.strFontSize = "20";
        pp.strItalic = "0";
        pp.strStrikethrough = "0";
        pp.strText = "";

        pp.strIsUnderline = "0";
        pp.strUnderlineClr = "Automatic";
        pp.strUnderlineType = "none";
        pp.strUnderlineClrThemeClr = "none";
        pp.strUnderlineClrThemeShade = "none";
        pp.strUnderlineClrThemeTint = "none";

        pp.strPagination = "0";

        pp.strCaps = "0";
        pp.strDoubleStrike = "0";
        pp.strEmboss = "0";
        pp.strEngrave = "0";
        pp.strHidden = "0";
        pp.strOutline = "0";
        pp.strShadow = "0";
        pp.strSmallCaps = "0";
        pp.strVAlign = "0";
        pp.strStyle = "none";
        pp.strHighlightColor = "No Color";
        pp.strSpacing = "0";
        pp.strWidth = "0";
        pp.strKerning = "0";
        pp.strPosition = "0";
        pp.strFootnoteID = "none";
        pp.strEndnoteID = "none";
        //added by neeraj
        pp.strBreakType = "none";
        pp.strAutoUpdate = "0";
        pp.strStyleType = "none";

        //shadow information
        pp.strShdowblur = "0";
        pp.strShdowdist = "0";
        pp.strShdowdir = "0";
        pp.strShdowsx = "0";
        pp.strShdowsy = "0";
        pp.strShdowSchemeclr = "none";
        pp.strShdowLummod = "none";
        pp.strShdowLumOff = "none";

        //ligature information
        pp.strLigature = "none";

        this.GetDefaultParaAttributes(pp);

        //For table paragraph parse table style also
        if (this.m_bParaInCell) {
          this.GetTableParaPropdNodes(xmlEle, pp);
        }
      }

      const runEle = xmlList[i];
      pp.strText = this.ExtractParaTextFromRNode(runEle);

      //pagination
      if (select("w:lastRenderedPageBreak", runEle, true) != null) {
        this.m_nPage++;
        pp.strPagination = this.m_nPage.toString();
      } else {
        pp.strPagination = this.m_nPage.toString();
      }

      xmlRPrEle = select("w:rPr", runEle, true) as Element;

      //check if the style is applied to the paragraph
      const strXPath =
        "/w:styles/w:style[@w:type='paragraph' and @w:styleId='" +
        this.m_ParaStyle +
        "']";
      if (!this.m_myPackageReader) {
        throw new Error("package not initialized!!");
      }
      const xmlStyleNode = select(
        strXPath,
        this.m_myPackageReader.partFileMap.get("m_StylesDoc") as Node,
        true
      ) as Element;

      if (xmlStyleNode != null) {
        this.FillRowParaPropsFromStyle(xmlStyleNode, pp);
      }
      //first FillParaAttrCollection from the styles node
      if (xmlRPrEle != null) {
        let strStyleName = WordUtils.extractStringFromVal(
          xmlRPrEle,
          "w:rStyle",
          select
        );
        if (strStyleName != "") {
          this.fillRowPropsStructureFromBaseFirst(strStyleName, pp);
        }
      }
      //default style

      //now fillfromthenode in the document.xml
      this.extractPropsFromRowPrNode(xmlRPrEle, pp);

      const xBR = select("w:br", runEle, true) as Element;
      if (xBR != null) {
        pp.strBreakType =
          xBR.getAttribute("w:type") == ""
            ? "shift-enter"
            : (xBR.getAttribute("w:type") as string);
      }
      if (nCount == 1 && i == 0) {
        this.m_paraAttrColl[0] = pp;
      } else {
        this.m_paraAttrColl.push(pp);
      }
    }
  };

  protected GetDefaultParaAttributes = (pp: paraFontAttributes): void => {
    if (!this.m_myPackageReader) {
      throw new Error("Package not initialized!!");
    }

    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });

    let strXPath = "/w:styles/w:docDefaults";
    const m_StylesDoc = this.m_myPackageReader.partFileMap.get(
      "m_StylesDoc"
    ) as XMLDocument;

    let xmlDefaultEle = select(strXPath, m_StylesDoc, true) as Element;

    // We need to extract the style in the table node
    if (xmlDefaultEle) {
      this.extractAttributesFromDefault(pp, xmlDefaultEle);
    }

    // Now check if the w:style node exist for paragraph with the style id as "Normal"
    strXPath = "/w:styles/w:style[@w:type='paragraph' and @w:styleId='Normal']";
    let xNextDefaultEle = select(strXPath, m_StylesDoc, true) as Element;

    if (xNextDefaultEle) {
      this.extractAttributesFromDefault(pp, xNextDefaultEle);
    }
    this.extractAttributesFromDefaultStyle(pp);
  };
  protected extractAttributesFromDefault = (
    pp: paraFontAttributes,
    xmlDefaultEle: Element
  ) => {
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });

    const rowProps = select(
      "descendant::w:rPr",
      xmlDefaultEle,
      true
    ) as Element;
    if (rowProps) {
      const fontName = select("w:rFonts", rowProps, true) as Element;
      if (fontName) {
        pp.strFontName =
          fontName.getAttribute("w:hAnsi") !== ""
            ? (fontName.getAttribute("w:hAnsi") as string)
            : (this.m_strThemeMinorFont as string);
      }

      pp.strFontSize =
        WordUtils.extractStringFromVal(rowProps, "w:sz", select) === ""
          ? pp.strFontSize
          : WordUtils.extractStringFromVal(rowProps, "w:sz", select);
    }
  };

  protected extractAttributesFromDefaultStyle = (
    pp: paraFontAttributes
  ): void => {
    if (!this.m_myPackageReader) {
      throw new Error("Package not initialized!!");
    }

    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });

    let strXPath = "/w:styles/w:docDefaults";
    let xmlDefaultEle = select(
      strXPath,
      this.m_myPackageReader.partFileMap.get("m_StylesDoc") as Node,
      true
    ) as Element;

    let xmlRPropsEle: Element | null = select(
      "w:rPrDefault/w:rPr",
      xmlDefaultEle,
      true
    ) as Element;
    this.extractPropsFromRowPrNode(xmlRPropsEle, pp);

    strXPath = "/w:styles/w:style[@w:type='paragraph' and @w:styleId='Normal']";
    let xNextDefaultEle = select(
      strXPath,
      this.m_myPackageReader.partFileMap.get("m_StylesDoc") as Node,
      true
    ) as Element;

    xmlRPropsEle = xNextDefaultEle
      ? (select("w:rPr", xNextDefaultEle, true) as Element)
      : null;
    if (xmlRPropsEle) {
      this.extractPropsFromRowPrNode(xmlRPropsEle, pp);
    }
  };

  protected extractPropsFromRowPrNode = (
    xmlRPropsEle: Element,
    pp: paraFontAttributes
  ): void => {
    let attrColl = null;
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    });

    if (xmlRPropsEle) {
      let tempFontEle = select("w:rFonts", xmlRPropsEle, true) as Element;

      if (tempFontEle) {
        attrColl = tempFontEle.attributes;

        let fontNode = attrColl.getNamedItem("w:hAnsi");

        if (fontNode) {
          pp.strFontName = fontNode.value;
        } else {
          fontNode = attrColl.getNamedItem("w:hAnsiTheme");
          if (fontNode != null) {
            const sFontTest = fontNode.textContent;
            switch (sFontTest) {
              case "minorHAnsi":
                pp.strFontName = this.m_strThemeMinorFont || "";
                break;

              case "majorHAnsi":
                pp.strFontName = this.m_strThemeMajorFont || "";
                break;
            }
          }
        }
      }

      pp.strFontSize =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:sz", select) === ""
          ? pp.strFontSize
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:sz", select);
      pp.strSpacing =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:spacing", select) === ""
          ? pp.strSpacing
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:spacing", select);
      pp.strKerning =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:kern", select) === ""
          ? pp.strKerning
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:kern", select);
      pp.strPosition =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:position", select) ===
        ""
          ? pp.strPosition
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:position", select);

      pp.strWidth =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:w", select) === ""
          ? pp.strWidth
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:w", select);

      let atemp = false;
      let btemp = false;
      if (pp.strBold == "1") {
        atemp = true;
      }
      if (WordUtils.extractString(xmlRPropsEle, "w:b", select) == "1") {
        btemp = true;
      }
      if (atemp !== btemp) {
        pp.strBold = "1";
      } else {
        pp.strBold = "0";
      }
      pp.strBold =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:b", select) == ""
          ? pp.strBold
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:b", select);

      atemp = false;
      btemp = false;
      if (pp.strItalic == "1") {
        atemp = true;
      }
      if (WordUtils.extractString(xmlRPropsEle, "w:i", select) == "1") {
        btemp = true;
      }
      if (atemp !== btemp) {
        pp.strItalic = "1";
      } else {
        pp.strItalic = "0";
      }
      pp.strItalic =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:i", select) == ""
          ? pp.strItalic
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:i", select);

      pp.strStrikethrough =
        WordUtils.extractString(xmlRPropsEle, "w:strike", select) == ""
          ? pp.strStrikethrough
          : WordUtils.extractString(xmlRPropsEle, "w:strike", select);
      pp.strStrikethrough =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:strike", select) == ""
          ? pp.strStrikethrough
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:strike", select);

      pp.strDoubleStrike =
        WordUtils.extractString(xmlRPropsEle, "w:dstrike", select) == ""
          ? pp.strDoubleStrike
          : WordUtils.extractString(xmlRPropsEle, "w:dstrike", select);
      pp.strDoubleStrike =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:dstrike", select) == ""
          ? pp.strDoubleStrike
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:dstrike", select);

      pp.strOutline =
        WordUtils.extractString(xmlRPropsEle, "w:outline", select) == ""
          ? pp.strOutline
          : WordUtils.extractString(xmlRPropsEle, "w:outline", select);
      pp.strOutline =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:outline", select) == ""
          ? pp.strOutline
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:outline", select);

      pp.strCaps =
        WordUtils.extractString(xmlRPropsEle, "w:caps", select) == ""
          ? pp.strCaps
          : WordUtils.extractString(xmlRPropsEle, "w:caps", select);
      pp.strCaps =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:caps", select) == ""
          ? pp.strCaps
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:caps", select);

      pp.strSmallCaps =
        WordUtils.extractString(xmlRPropsEle, "w:smallCaps", select) == ""
          ? pp.strSmallCaps
          : WordUtils.extractString(xmlRPropsEle, "w:smallCaps", select);
      pp.strSmallCaps =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:smallCaps", select) ==
        ""
          ? pp.strSmallCaps
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:smallCaps", select);

      pp.strEmboss =
        WordUtils.extractString(xmlRPropsEle, "w:emboss", select) == ""
          ? pp.strEmboss
          : WordUtils.extractString(xmlRPropsEle, "w:emboss", select);
      pp.strEmboss =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:emboss", select) == ""
          ? pp.strEmboss
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:emboss", select);

      pp.strEngrave =
        WordUtils.extractString(xmlRPropsEle, "w:imprint", select) == ""
          ? pp.strEngrave
          : WordUtils.extractString(xmlRPropsEle, "w:imprint", select);
      pp.strEngrave =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:imprint", select) == ""
          ? pp.strEngrave
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:imprint", select);

      pp.strHidden =
        WordUtils.extractString(xmlRPropsEle, "w:vanish", select) == ""
          ? pp.strHidden
          : WordUtils.extractString(xmlRPropsEle, "w:vanish", select);
      pp.strHidden =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:vanish", select) == ""
          ? pp.strHidden
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:vanish", select);

      pp.strHighlightColor =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:highlight", select) ==
        ""
          ? pp.strHighlightColor
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:highlight", select);

      pp.strLigature =
        WordUtils.extractString(xmlRPropsEle, "w14:ligatures", select) == ""
          ? pp.strLigature
          : WordUtils.extractString(xmlRPropsEle, "w14:ligatures", select);
      pp.strLigature =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w14:ligatures", select) ==
        ""
          ? pp.strLigature
          : WordUtils.extractStringFromVal(
              xmlRPropsEle,
              "w14:ligatures",
              select
            );
      this.RetrieveExtraColorInfo(xmlRPropsEle, pp);

      // Read textoutline information
      this.RetrieveTextOutlineInfo(xmlRPropsEle, pp);

      //Read shadow information
      this.RetrieveShadowInformation(xmlRPropsEle, pp);

      ////Read Ligature information
      //RetrieveLigatureInformation(xmlRPropsEle, ref pp);

      this.RetrieveUnderlineInfo(xmlRPropsEle, pp);
      pp.strVAlign =
        WordUtils.extractStringFromVal(xmlRPropsEle, "w:vertAlign", select) ==
        ""
          ? pp.strVAlign
          : WordUtils.extractStringFromVal(xmlRPropsEle, "w:vertAlign", select);
    }

    if (this.m_xmlRunEle) {
      const xFootnoteRef = select(
        "w:footnoteReference",
        this.m_xmlRunEle as Node,
        true
      ) as Element;
      let strFootnoteRef = "";
      if (xFootnoteRef != null) {
        strFootnoteRef = xFootnoteRef.getAttribute("w:id") || "";
      }
      pp.strFootnoteID =
        strFootnoteRef == "" ? pp.strFootnoteID : strFootnoteRef;

      const xEndnoteRef = select(
        "w:endnoteReference",
        this.m_xmlRunEle as Node,
        true
      ) as Element;
      let strEndnoteRef = "";
      if (xEndnoteRef) {
        strEndnoteRef = xEndnoteRef.getAttribute("w:id") || "";
      }
      pp.strEndnoteID = strEndnoteRef == "" ? pp.strEndnoteID : strEndnoteRef;
    }
    this.RetrieveShadingInformation(xmlRPropsEle, pp);
  };

  protected RetrieveExtraColorInfo = (
    xmlRPropsEle: Element,
    pp: paraFontAttributes
  ): void => {
    let attrColl = null;
    let strValue = "";
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    });
    pp.strColor =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:color", select) == ""
        ? pp.strColor
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:color", select);
    strValue = pp.strColor;

    const xGradNode = select(
      "w14:textFill/w14:gradFill",
      xmlRPropsEle,
      true
    ) as Element;

    if (xGradNode == null) {
      const xmlTemp = select("w:color", xmlRPropsEle, true) as Element;

      if (xmlTemp) {
        pp.strType = "solid";
        attrColl = xmlTemp.attributes;
        const themecolorNode = attrColl.getNamedItem("w:themeColor");
        const themeshadeNode = attrColl.getNamedItem("w:themeShade");
        const themetintNode = attrColl.getNamedItem("w:themeTint");
        if (themecolorNode != null) {
          pp.strThemeColor = themecolorNode.value;
          if (themeshadeNode != null) {
            pp.strThemeShade = themeshadeNode.value;
          }
          if (themetintNode != null) {
            pp.strThemeTint = themetintNode.value;
          }
        }
      }

      const xmlSolidFill = select(
        "w14:textFill/w14:solidFill",
        xmlRPropsEle,
        true
      ) as Element;

      if (xmlSolidFill) {
        pp.strType = "solid";
        let xSchemClr = select("w14:schemeClr", xmlSolidFill) as Element;
        if (xSchemClr != null) {
          pp.strThemeColor = xSchemClr.getAttribute("val") || "";

          const xTint = select("w14:tint", xSchemClr, true) as Element;
          if (xTint) {
            pp.strThemeTint = xTint.getAttribute("val") || "";
          }
          const xShade = select("w14:shade", xSchemClr, true) as Element;
          if (xShade) {
            pp.strThemeShade = xShade.getAttribute("val") || "";
          }
        }
      }

      if (xmlTemp == null && xmlSolidFill == null) {
        pp.strType = "nofill";
      }
    } else {
      // pp.strType = "gradiant";
      // const gradStopList = select("descendant::w14:gs", xGradNode) as Element[];
      //     pp.gradientStops = new GradientStopAttributes[gradStopList.Count];
      //     let i = 0, tempNode;
      //     gradStopList.forEach(element => {
      //       pp.gradientStops[i].strPos = element.getAttribute("w14:pos") || '';
      //         pp.gradientStops[i].strTClr = element.getAttribute("w14:val") || '';
      //         tempNode = select("w14:tint", element, true) as Element;
      //         if (tempNode != null)
      //             pp.gradientStops[i].strTTint = tempNode.getAttribute("w14:val") || '';
      //         tempNode = select("w14:satMod", element.childNodes[0], true) as Element;
      //         if (tempNode != null)
      //             pp.gradientStops[i].strTSat = tempNode.getAttribute("w14:val") || '';
      //         i++;
      //     });
    }
  };

  protected RetrieveTextOutlineInfo = (
    xmlRPropsEle: Element,
    pp: paraFontAttributes
  ): void => {
    const select = xpath.useNamespaces({
      w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    });

    const xTextOutlineNode = select(
      "w14:textOutline",
      xmlRPropsEle,
      true
    ) as Element;
    if (xTextOutlineNode != null) {
      const xmlSchemeColorNode = select(
        "w14:solidFill/w14:schemeClr",
        xTextOutlineNode,
        true
      ) as Element;
      if (xmlSchemeColorNode) {
        pp.strTextOutlineThemColor =
          xmlSchemeColorNode.getAttribute("w14:val") || "";

        let xmlTempNode = select(
          "w14:shade",
          xmlSchemeColorNode,
          true
        ) as Element;
        if (xmlTempNode) {
          pp.strTextOutlineThemeShade =
            xmlTempNode.getAttribute("w14:val") || "";
        }
        xmlTempNode = select("w14:satMod", xmlSchemeColorNode, true) as Element;
        if (xmlTempNode) {
          pp.strTextOutlineSatMode = xmlTempNode.getAttribute("w14:val") || "";
        }
      }

      const xmlLineTypeNode = select(
        "w14:prstDash",
        xTextOutlineNode,
        true
      ) as Element;
      if (xmlLineTypeNode) {
        pp.strTextOutlineLineType =
          xmlLineTypeNode.getAttribute("w14:val") || "";
      }

      pp.strTextOutlineWidth = xTextOutlineNode.getAttribute("w14:w") || "";
    }
  };

  protected RetrieveShadowInformation = (
    xmlRPropsEle: Element,
    pp: paraFontAttributes
  ): void => {
    const select = xpath.useNamespaces({
      w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    });

    const xShadowNode = select("w14:shadow", xmlRPropsEle, true) as Element;
    if (xShadowNode) {
      pp.strShdowblur = xShadowNode.getAttribute("w14:blurRad") || "";
      pp.strShdowdist = xShadowNode.getAttribute("w14:dist") || "";
      pp.strShdowdir = xShadowNode.getAttribute("w14:dir") || "";
      pp.strShdowsx = xShadowNode.getAttribute("w14:sx") || "";
      pp.strShdowsy = xShadowNode.getAttribute("w14:sy") || "";

      let xmlSchemeColorNode = select(
        "w14:schemeClr",
        xShadowNode,
        true
      ) as Element;
      if (xmlSchemeColorNode != null) {
        pp.strShdowSchemeclr = xmlSchemeColorNode.getAttribute("w14:val") || "";

        let xmlTempNode = select(
          "w14:alpha",
          xmlSchemeColorNode,
          true
        ) as Element;
        if (xmlTempNode) {
          pp.strShdowLummod = xmlTempNode.getAttribute("w14:val") || "";
        }

        xmlTempNode = select("w14:lumMod", xmlSchemeColorNode, true) as Element;
        if (xmlTempNode) {
          pp.strShdowLumOff = xmlTempNode.getAttribute("w14:val") || "";
        }
      }
    }
  };

  protected RetrieveUnderlineInfo = (
    xmlRPropsEle: Element,
    pp: paraFontAttributes
  ): void => {
    let attrColl = null;
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    });
    const tempEle = select("w:u", xmlRPropsEle, true) as Element;
    if (tempEle) {
      pp.strIsUnderline = "0";
      attrColl = tempEle.attributes;
      const underlinetypeNode = attrColl.getNamedItem("w:val");
      const underlinecolorNode = attrColl.getNamedItem("w:color");
      const underlinethemecolorNode = attrColl.getNamedItem("w:themeColor");
      const underlinethemeshadeNode = attrColl.getNamedItem("w:themeShade");
      const underlinethemetintNode = attrColl.getNamedItem("w:themeTint");

      if (underlinetypeNode != null) {
        pp.strUnderlineType = underlinetypeNode.value;
        pp.strIsUnderline = "1";
      }

      if (underlinecolorNode != null) {
        pp.strUnderlineClr = underlinecolorNode.value;
      } else if (pp.strUnderlineClr.length <= 0) {
        pp.strUnderlineClr = "Automatic";
      }

      if (underlinethemecolorNode != null) {
        pp.strUnderlineClrThemeClr = underlinethemecolorNode.value;

        if (underlinethemeshadeNode != null) {
          pp.strUnderlineClrThemeShade = underlinethemeshadeNode.value;
        }

        if (underlinethemetintNode != null) {
          pp.strUnderlineClrThemeTint = underlinethemetintNode.value;
        }
      }
    }
  };

  protected RetrieveShadingInformation = (
    xmlRPropsEle: Element,
    pp: paraFontAttributes
  ): void => {
    if (xmlRPropsEle != null) {
      let strValue = "";
      let strPattValue = "";
      const select = xpath.useNamespaces({
        w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      });
      const xmlTemp = select("w:shd", xmlRPropsEle, true) as Element;
      if (xmlTemp) {
        pp.strShdPattCol =
          WordUtils.ExtractAttribute(xmlTemp, "w:color") == ""
            ? pp.strShdPattCol
            : WordUtils.ExtractAttribute(xmlTemp, "w:color");
        pp.strShdPattStyle =
          WordUtils.ExtractAttribute(xmlTemp, "w:val") == ""
            ? pp.strShdPattStyle
            : WordUtils.ExtractAttribute(xmlTemp, "w:val");
        pp.strShdPattThemeClr =
          WordUtils.ExtractAttribute(xmlTemp, "w:themeColor") == ""
            ? pp.strShdPattThemeClr
            : WordUtils.ExtractAttribute(xmlTemp, "w:themeColor");
        pp.strShdPattThemeShade =
          WordUtils.ExtractAttribute(xmlTemp, "w:themeShade") == ""
            ? pp.strShdPattThemeShade
            : WordUtils.ExtractAttribute(xmlTemp, "w:themeShade");
        pp.strShdPattThemeTint =
          WordUtils.ExtractAttribute(xmlTemp, "w:themeTint") == ""
            ? pp.strShdPattThemeTint
            : WordUtils.ExtractAttribute(xmlTemp, "w:themeTint");
        pp.strShdColor =
          WordUtils.ExtractAttribute(xmlTemp, "w:fill") == ""
            ? pp.strShdColor
            : WordUtils.ExtractAttribute(xmlTemp, "w:fill");
        pp.strShdThemeColor =
          WordUtils.ExtractAttribute(xmlTemp, "w:themeFill") == ""
            ? pp.strShdThemeColor
            : WordUtils.ExtractAttribute(xmlTemp, "w:themeFill");
        pp.strShdThemeShade =
          WordUtils.ExtractAttribute(xmlTemp, "w:themeFillShade") == ""
            ? pp.strShdThemeShade
            : WordUtils.ExtractAttribute(xmlTemp, "w:themeFillShade");
        pp.strShdThemeTint =
          WordUtils.ExtractAttribute(xmlTemp, "w:themeFillTint") == ""
            ? pp.strShdThemeTint
            : WordUtils.ExtractAttribute(xmlTemp, "w:themeFillTint");

        strValue = pp.strShdColor;
        strPattValue = pp.strShdPattCol;

        if (pp.strShdPattStyle == "clear" || pp.strShdPattStyle == "none") {
          pp.strShdPattStyle = "Clear";
        }

        if (pp.strShdPattCol == "auto" || pp.strShdPattCol == "none") {
          pp.strShdPattCol = "Automatic";
        }

        if (pp.strShdColor == "auto" || pp.strShdColor == "none") {
          pp.strShdColor = "No Color";
        }
      }
    }
  };

  protected FillRowParaPropsFromStyle(
    xmlCStyleNode: Element,
    pp: paraFontAttributes
  ) {
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });
    pp.strStyle = WordUtils.extractStringFromVal(
      xmlCStyleNode,
      "w:name",
      select
    );

    const xautoRedefine = select("w:autoRedefine", xmlCStyleNode, true);
    if (xautoRedefine != null) {
      pp.strAutoUpdate = "1";
    }
    if (xmlCStyleNode.getAttribute("w:type")) {
      pp.strStyleType = xmlCStyleNode.getAttribute("w:type") || "";

      // Updating code to remove trailing char from "Linked Paragraph and Character" type custom styles....MYITLAB-15687
      const linkNode = select("w:link", xmlCStyleNode, true);
      if (
        pp.strStyleType.toLowerCase() == "character" &&
        linkNode != null &&
        pp.strStyle.toLowerCase().endsWith(" char")
      ) {
        pp.strStyle = pp.strStyle.substring(0, pp.strStyle.lastIndexOf(" ")); // need to check this. - vipul
      }
    }
    const xmlRPropsEle = select("w:rPr", xmlCStyleNode, true) as Element;

    let strBasedOn = WordUtils.extractStringFromVal(
      xmlCStyleNode,
      "w:basedOn",
      select
    );

    if (strBasedOn != "") {
      this.FillPropsStructureFromBaseFirst(strBasedOn, pp);
    }
    this.extractPropsFromRowPrNode(xmlRPropsEle, pp);
  }

  protected FillPropsStructureFromBaseFirst = (
    strBasedOn: string,
    pp: paraFontAttributes
  ) => {
    if (!this.m_myPackageReader) {
      throw new Error("Package not initialized!!");
    }
    const strXPath =
      "/w:styles/w:style[@w:type='paragraph' and @w:styleId='" +
      strBasedOn +
      "']";
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });
    const xmlPStyleNode = select(
      strXPath,
      this.m_myPackageReader.partFileMap.get("m_StylesDoc") as Node,
      true
    ) as Element;
    if (xmlPStyleNode == null) {
      return;
    }

    let strBasedOnNew = WordUtils.extractStringFromVal(
      xmlPStyleNode,
      "w:basedOn",
      select
    );

    if (strBasedOnNew != "") {
      this.FillPropsStructureFromBaseFirst(strBasedOnNew, pp);
    }
    this.FillParaPropsFromStyle(xmlPStyleNode, pp);
  };

  protected FillParaPropsFromStyle = (
    xmlPStyleNode: Element,
    pp: paraFontAttributes
  ): void => {
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });
    const xmlRPropsEle = select("w:rPr", xmlPStyleNode, true) as Element;
    if (xmlRPropsEle != null) {
      this.extractPropsFromRowPrNode(xmlRPropsEle, pp);
    }
  };

  protected GetTableParaPropdNodes = (
    XmlOffice12Node: Element,
    pp: paraFontAttributes
  ): void => {
    //First Check if the pPr node is present in the table style
    const xmlTStyleNode = this.ReturnTStyleNode();
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });
    if (xmlTStyleNode) {
      const xmlpPrEle = select("w:pPr", xmlTStyleNode, true) as Element;
      if (xmlpPrEle) {
        this.extractPropsFromRowPrNode(xmlpPrEle, pp);
      }

      //Then Check if the rPr node is present in the table style
      const xmlrPrEle = select("w:rPr", xmlTStyleNode, true) as Element;
      if (xmlrPrEle) {
        this.extractPropsFromRowPrNode(xmlrPrEle, pp);
      }
    }

    //Fisrt check if the CNF node is present in the pPr node
    const xmlPProps = select("w:pPr", XmlOffice12Node, true) as Element;

    let strCnfStyleFrompPr = WordUtils.extractStringFromVal(
      xmlPProps,
      "w:cnfStyle",
      select
    );

    if (strCnfStyleFrompPr != "") {
      const xmlTblStyleNode = this.ExtractTableStyleNode(strCnfStyleFrompPr);
      if (xmlTblStyleNode) {
        const xmlrPrEle = select("w:rPr", xmlTblStyleNode, true) as Element;
        if (xmlrPrEle != null) {
          this.extractPropsFromRowPrNode(xmlrPrEle, pp);
        }
      }
    }

    //Now check if the Cnf is presnet in rPr node
    let xCnfStyle = select(
      "parent::node()/w:tcPr/w:cnfStyle",
      XmlOffice12Node,
      true
    ) as Element;

    if (!xCnfStyle) {
      //Now check if the Cnf is presnet in trPr node
      xCnfStyle = select(
        "parent::node()/parent::node()/w:trPr/w:cnfStyle",
        XmlOffice12Node,
        true
      ) as Element;
      if (xCnfStyle == null) {
        return;
      }
    }

    let attrColl = xCnfStyle.attributes;
    if (attrColl) {
      let strCnfStyle =
        attrColl.getNamedItem("w:val") != null
          ? attrColl.getNamedItem("w:val")?.value
          : "";
      if (strCnfStyle != "") {
        const xmlTblStyleNode = this.ExtractTableStyleNode(strCnfStyle || "");
        if (xmlTblStyleNode != null) {
          const xmlrPrEle = select("w:rPr", xmlTblStyleNode, true) as Element;
          if (xmlrPrEle != null) {
            this.extractPropsFromRowPrNode(xmlrPrEle, pp);
          }
        }
      }
    }
  };

  protected ExtractParaTextFromRNode = (XmlRNode: Element): string => {
    let strParaText = "";

    let objNode = XmlRNode.firstChild as Element;
    while (objNode != null) {
      //XmlElement xmlTempELe = (XmlElement)nodeRChildList[j];

      switch (objNode.nodeName) {
        case "w:t":
          strParaText += objNode.textContent;
          break;
        case "w:fldSimple":
          strParaText += "@formula@";
          break;
        case "w:tab":
          strParaText += "\\t";
          break;
        case "w:noBreakHyphen":
          //strParaText += "NonBreakingHyphen";
          strParaText += "\x2011";
          break;
        case "w:softHyphen":
          //strParaText += "OptionalHyphen";
          strParaText += "\x00AD";
          break;

        case "w:sym":
          if (objNode.attributes != null) {
            const xChar = objNode.getAttribute("w:char");
            if (xChar) {
              let temp = xChar;
              let itemp = parseInt(temp, 16);
              strParaText += String.fromCharCode(itemp);
            }
          }
          break;
      }
      objNode = objNode.nextSibling as Element;
    }

    return strParaText;
  };

  protected fillRowPropsStructureFromBaseFirst = (
    strBasedOn: string,
    pp: paraFontAttributes
  ): void => {
    if (!this.m_myPackageReader) {
      throw new Error("Package not initialized!!");
    }
    const strXPath = `/w:styles/w:style[@w:type='character' and @w:styleId='${strBasedOn}']`;
    const xmlDoc: XMLDocument = this.m_myPackageReader.partFileMap.get(
      "m_StylesDoc"
    ) as XMLDocument;
    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });
    const xmlCStyleNode = select(strXPath, xmlDoc, true) as Element;

    if (!xmlCStyleNode) {
      return;
    }

    const strBasedOnNew = WordUtils.extractStringFromVal(
      xmlCStyleNode,
      "w:basedOn",
      select
    );

    if (strBasedOnNew !== "") {
      this.fillRowPropsStructureFromBaseFirst(strBasedOnNew, pp);
    }
    this.FillRowParaPropsFromStyle(xmlCStyleNode, pp);
  };
}
