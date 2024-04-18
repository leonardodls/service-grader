import { PackageReader } from "../../ReaderBase/PackageReader";
import xpath from "xpath";
import { CommonFunctions } from "../../Utils/common";
import { Utils } from "./Utils";
import { WordUtils } from "../../WordReader/CommonClasses/WordUtils";
import { WordUtils2010 } from "../../WordReader2010/CommonClasses/WordUtils";

export class TextAttributes2013 {
  public m_bParaInCell: boolean = false;
  private m_xmlParaEle: Element | null = null;
  private m_myXmlDoc: XMLDocument | null = null;
  private m_CMLProps: Element | null = null;
  private m_myPackageReader: PackageReader | null = null;
  protected m_strThemeMinorFont: string | null = "";
  protected m_strThemeMajorFont: string | null = "";
  protected m_docUri: string = "";
  protected m_paraAttrColl: [] = [];
  protected m_strShapeID: string | null = null;

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
  };

  GenerateParaText = (XmlOffice12Node: Element, paraProps: Element) => {
    let strParaText: string = "",
      nIndex: number = 0,
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
                const xFomulas = WordUtils.ReturnNode(
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
                strParaText,
                nIndex
              );

              sBuffer += sText;
              let xFrm = xpath.select1(
                `formulas/formula[@id='${nFormulaCount.toString()}']`,
                paraProps
              );
              console.log("paraProps: ", paraProps.textContent);
              // const xFrm = paraProps.getElementsByTagName(
              //   "formulas/formula[@id='" + nFormulaCount.toString() + "']"
              // )[0];
              if (xFrm)
                WordUtils.ReturnNode(
                  this.m_myXmlDoc as XMLDocument,
                  xFrm as Element,
                  "infotxt"
                ).textContent = sBuffer;
              break;
            case "end":
              if (bFormulaDone == false && nFormulaCount > 0) {
                const xFomulas = WordUtils.ReturnNode(
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
              strParaText,
              nIndex
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

  protected ExtractParaTextAndSpecialcharsFromRNode = (
    XmlRNode: Element,
    sParaTxt: string,
    nIndex: number
  ) => {
    let strParaText = "";
    const xSCs = WordUtils2010.ReturnNode(
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

          // Check for special chars in the run text
          const cs = p?.split("") as string[];
          const b: Buffer = Buffer.from(p as string, "utf16le"); // byte[] b = Encoding.Unicode.GetBytes(cs);

          const nLength = cs.length;
          let c: string[] = Array(nLength).fill(""); //char[] c = new char[nLength];
          let b1: number, b2: number, f: boolean;

          const decodedString = b.toString("utf16le"); //Encoding.Unicode.GetDecoder().Convert(b, 0, b.length, c, 0, nLength, true, out b1, out b2, out f);

          for (let i = 0; i < nLength; i++) {
            switch (c[i].codePointAt(0)) {
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
                  nIndex.toString()
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
                nIndex++;
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
                  nIndex.toString()
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
                nIndex++;
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
                  nIndex.toString()
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
                nIndex++;
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
                  nIndex.toString()
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
                nIndex++;
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
                  nIndex.toString()
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
                nIndex++;
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
            nIndex.toString()
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
          nIndex++;
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
            nIndex.toString()
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
          nIndex++;
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
                nIndex.toString()
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
              nIndex++;
            }
          }
          break;
      }
      objNode = objNode.nextSibling;
    }
    return strParaText;
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
}
