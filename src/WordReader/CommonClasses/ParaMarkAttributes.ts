import { PackageReader } from "../../ReaderBase/PackageReader";
import { CommonFunctions } from "../../Utils/common";
import { Utils } from "../../WordReader2013/CommonClasses/Utils";
import { WordUtils } from "./WordUtils";
import xpath from "xpath";

export class ParaMarkAttributes {
  protected m_xmlrPrEle: Element | null = null;
  protected m_myXmlDoc: XMLDocument | null = null;
  public m_myPackageReader: PackageReader | null = null;
  protected m_docUri: string = "";

  public parseParaMarkAttributes = async (
    XmlOffice12Node: Element,
    xmlDoc: XMLDocument,
    CMLNode: Element,
    CurrentPackageReader: any,
    strTagName: string
  ): Promise<void> => {
    this.m_xmlrPrEle = XmlOffice12Node;
    this.m_myXmlDoc = xmlDoc;
    this.m_myPackageReader = CurrentPackageReader;

    if (!this.m_myPackageReader) {
      throw new Error("Package reader not initialized!!");
    }

    let strDocUriString = await this.m_myPackageReader.ReturnBaseXML(
      null,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    );
    this.m_docUri = CommonFunctions.PrependStringToURIPath(
      strDocUriString,
      "/"
    );

    this.extractPropsFromrPrNode(this.m_xmlrPrEle, strTagName, CMLNode);
  };

  protected extractPropsFromrPrNode(
    xmlRPropsEle: Element,
    strTagName: string,
    CMLPProps: Element
  ): void {
    const paramark = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      CMLPProps,
      strTagName
    );
    this.getDefaultParaAttributes(paramark);
    if (xmlRPropsEle) {
      if (
        xmlRPropsEle.parentNode &&
        xmlRPropsEle.parentNode.nodeName === "w:style"
      ) {
        const attrColl = xmlRPropsEle.parentElement?.attributes;
        const styleId = attrColl?.getNamedItem("w:styleId")
          ? attrColl?.getNamedItem("w:styleId")?.value
          : "none";
        const elem = WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          paramark,
          "style"
        );
        elem.textContent = styleId || "";
      }
      const xRowStyle = xmlRPropsEle.getElementsByTagName("w:rStyle")[0];
      if (xRowStyle) {
        const select = xpath.useNamespaces({
          w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        });
        const strStyleName = WordUtils.extractStringFromVal(
          xmlRPropsEle,
          "w:rStyle",
          select
        );
        // this.extractRowPropertiesFromStyle(strStyleName, paramark);
        this.fillRowPropsStructureFromBaseFirst(strStyleName, paramark);
      }
    } else {
      WordUtils.returnNode(
        this.m_myXmlDoc as XMLDocument,
        paramark,
        "style"
      ).textContent = "none";
    }
    this.extractProperties(xmlRPropsEle, paramark);
  }

  protected getDefaultParaAttributes(paramark: Element): void {
    if (!this.m_myPackageReader) {
      throw new Error("Package reader not initialized!!");
    }
    // theme font
    const ns = {
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    };
    const select = xpath.useNamespaces(ns);

    let strXPath = "/w:styles/w:docDefaults";
    let xmlDefaultEle = select(
      strXPath,
      this.m_myPackageReader.partFileMap.get("m_StylesDoc") as Node,
      true
    ) as Element;

    // We need to extract the style in the table node
    if (xmlDefaultEle) {
      let rowProps = select(
        "w:rPrDefault/w:rPr",
        xmlDefaultEle,
        true
      ) as Element;
      if (rowProps) {
        this.getAttributesFromStyles(paramark, rowProps);
      }
    }

    // Now check if the w:style node exists for paragraph with the style id as "Normal"
    strXPath = "/w:styles/w:style[@w:type='paragraph' and @w:styleId='Normal']";
    let xNextDefaultEle = select(
      strXPath,
      this.m_myPackageReader.partFileMap.get("m_StylesDoc") as Node,
      true
    ) as Element;

    if (xNextDefaultEle) {
      let rowProps = select("w:rPr", xNextDefaultEle, true) as Element;
      if (rowProps) {
        this.getAttributesFromStyles(paramark, rowProps);
      }
    }
  }

  protected getAttributesFromStyles(paramark: Element, xmlRPropsEle: Element) {
    // font name
    const ns = {
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    };
    const select = xpath.useNamespaces(ns);
    const fontname = select("w:rFonts", xmlRPropsEle, true) as Element;
    if (fontname) {
      const attrColl = fontname.attributes;
      const hAnsiAttr = attrColl.getNamedItem("w:hAnsi");
      WordUtils.returnNode(
        this.m_myXmlDoc as XMLDocument,
        paramark,
        "name"
      ).textContent = hAnsiAttr
        ? hAnsiAttr.value
        : WordUtils.returnNode(this.m_myXmlDoc as XMLDocument, paramark, "name")
            .textContent;
    }
    // other attributes similar to the pattern above
    const attributes = [
      "sz",
      "b",
      "i",
      "smallCaps",
      "caps",
      "strike",
      "outline",
      "dstrike",
      "hidden",
      "engrave",
      "emboss",
      "valign",
      "style",
      "spac",
      "wd",
      "kern",
      "pos",
    ];
    attributes.forEach((attr) => {
      const value = WordUtils.extractStringFromVal(
        xmlRPropsEle,
        `w:${attr}`,
        select
      );
      WordUtils.returnNode(
        this.m_myXmlDoc as XMLDocument,
        paramark,
        attr
      ).textContent =
        value !== ""
          ? value
          : WordUtils.returnNode(this.m_myXmlDoc as XMLDocument, paramark, attr)
              .textContent;
    });

    // Special handling for Underline and Color nodes
    this.generateUnderlineNode(xmlRPropsEle, paramark);
    this.generateColorNode(xmlRPropsEle, paramark);
    this.generateShadowNode(xmlRPropsEle, paramark);
  }

  protected generateUnderlineNode(
    xmlRPropsEle: Element,
    paraMark: Element
  ): void {
    let underline = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paraMark,
      "u"
    );
    let xmlUEle: Element | null = null;

    if (xmlRPropsEle) {
      xmlUEle = xmlRPropsEle.getElementsByTagName("w:u")[0] || null;

      if (xmlUEle) {
        let attrColl = xmlUEle.attributes;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "value"
        ).textContent = attrColl.getNamedItem("w:val")
          ? "1"
          : WordUtils.returnNode(
              this.m_myXmlDoc as XMLDocument,
              underline,
              "value"
            ).textContent;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "type"
        ).textContent = attrColl.getNamedItem("w:val")
          ? attrColl.getNamedItem("w:val")!.value
          : WordUtils.returnNode(
              this.m_myXmlDoc as XMLDocument,
              underline,
              "type"
            ).textContent;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument as XMLDocument,
          underline,
          "clr"
        ).textContent = attrColl.getNamedItem("w:color")
          ? attrColl.getNamedItem("w:color")!.value
          : "Automatic";
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "themeclr"
        ).textContent = attrColl.getNamedItem("w:themeColor")
          ? attrColl.getNamedItem("w:themeColor")!.value
          : WordUtils.returnNode(
              this.m_myXmlDoc as XMLDocument,
              underline,
              "themeclr"
            ).textContent;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "themeshade"
        ).textContent = attrColl.getNamedItem("w:themeShade")
          ? attrColl.getNamedItem("w:themeShade")!.value
          : WordUtils.returnNode(
              this.m_myXmlDoc as XMLDocument,
              underline,
              "themeshade"
            ).textContent;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "themetint"
        ).textContent = attrColl.getNamedItem("w:themeTint")
          ? attrColl.getNamedItem("w:themeTint")!.value
          : WordUtils.returnNode(
              this.m_myXmlDoc as XMLDocument,
              underline,
              "themetint"
            ).textContent;
      } else {
        const nodesToBeChecked = [
          "value",
          "type",
          "clr",
          "themeclr",
          "themeshade",
          "themetint",
        ];
        nodesToBeChecked.forEach((nodeToBeChecked) => {
          WordUtils.returnNode(
            this.m_myXmlDoc as XMLDocument,
            underline,
            nodeToBeChecked
          );
        });
      }
    } else {
      const nodesToBeChecked = [
        "value",
        "type",
        "clr",
        "themeclr",
        "themeshade",
        "themetint",
      ];
      nodesToBeChecked.forEach((nodeToBeChecked) => {
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          nodeToBeChecked
        );
      });
    }
  }

  protected generateColorNode = (
    xmlRPropsEle: Element,
    paramark: Element
  ): void => {
    const color = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "clr"
    );

    let xType = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      color,
      "type"
    );
    const xHex = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      color,
      "hex"
    );
    const xThmClr = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      color,
      "themeclr"
    );
    const xThmShd = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      color,
      "themeshade"
    );
    const xThmTint = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      color,
      "themetint"
    );

    if (xmlRPropsEle) {
      const select = xpath.useNamespaces({
        w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        w14: "http://schemas.microsoft.com/office/word/2010/wordml",
      });
      // Select gradient node
      let xTempNode = select(
        "w14:textFill/w14:gradFill",
        xmlRPropsEle,
        true
      ) as Element | null;

      // If no gradient node (CR fix - TFS: 386780)
      if (!xTempNode) {
        let xmlUEle = select("w:color", xmlRPropsEle, true) as Element | null;
        if (xmlUEle) {
          xType.textContent = "solid";
          let attrColl = xmlUEle.attributes;
          xHex.textContent =
            WordUtils.extractStringFromVal(xmlRPropsEle, "w:color", select) ==
            ""
              ? xHex.textContent
              : WordUtils.extractStringFromVal(xmlRPropsEle, "w:color", select);
          xThmClr.textContent =
            attrColl.getNamedItem("w:themeColor")?.value || xThmClr.textContent;
          xThmShd.textContent =
            attrColl.getNamedItem("w:themeShade")?.value || xThmShd.textContent;
          xThmTint.textContent =
            attrColl.getNamedItem("w:themeTint")?.value || xThmTint.textContent;
        } else {
          // Add other cases here - if any
          xType.textContent = "nofill";
        }
      } else {
        color.childNodes.forEach((child) => color.removeChild(child));
        xType = Utils.CreateNode(
          this.m_myXmlDoc as XMLDocument,
          color,
          "type",
          "gradient"
        );
        const xStops = Utils.CreateNode(
          this.m_myXmlDoc as XMLDocument,
          color,
          "stops",
          ""
        );
        const gradStopList = select(
          "descendant::w14:gs",
          xTempNode
        ) as Element[];
        let tempNode;

        if (gradStopList) {
          for (let i = 0; i < gradStopList.length; i++) {
            let xStop = Utils.CreateNode(
              this.m_myXmlDoc as XMLDocument,
              xStops,
              "stop",
              ""
            );
            Utils.CreateAttribute(
              this.m_myXmlDoc as XMLDocument,
              xStop,
              "pos",
              gradStopList[i].getAttribute("w14:pos") || ""
            );

            let clrValue = gradStopList[i].children[0].getAttribute("w14:val");
            Utils.CreateNode(
              this.m_myXmlDoc as XMLDocument,
              xStop,
              "clr",
              clrValue || ""
            );

            let tempNode =
              gradStopList[i].children[0].getElementsByTagName("w14:tint")[0];
            if (tempNode) {
              Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                xStop,
                "tint",
                tempNode.getAttribute("w14:val") || ""
              );
            } else {
              Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                xStop,
                "tint",
                ""
              );
            }

            tempNode =
              gradStopList[i].children[0].getElementsByTagName("w14:satMod")[0];
            if (tempNode) {
              Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                xStop,
                "sat",
                tempNode.getAttribute("w14:val") || ""
              );
            } else {
              Utils.CreateNode(
                this.m_myXmlDoc as XMLDocument,
                xStop,
                "sat",
                ""
              );
            }
          }
        }
      }
    } else {
      xType.textContent = "nofill";
      xHex.textContent = "Automatic";
      xThmClr.textContent = "none";
      xThmShd.textContent = "none";
      xThmTint.textContent = "none";
    }
  };

  protected generateShadowNode = (
    xmlRPropsEle: Element,
    paramark: Element
  ): void => {
    const shadow = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "shadow"
    );
    const xblur = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      shadow,
      "blur"
    );
    const xdist = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      shadow,
      "dist"
    );
    const xdir = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      shadow,
      "angle"
    );
    const xsx = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      shadow,
      "sizex"
    );
    const xsy = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      shadow,
      "sizey"
    );
    const xschemeclr = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      shadow,
      "schemeclr"
    );
    const xlummod = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      shadow,
      "lummod"
    );
    const xlumoff = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      shadow,
      "lumoff"
    );

    let xmlUEle = null;

    if (xmlRPropsEle) {
      const select = xpath.useNamespaces({
        w14: "http://schemas.microsoft.com/office/word/2010/wordml",
      });

      xmlUEle = select("w14:shadow", xmlRPropsEle, true) as Element;
      if (xmlUEle) {
        const attrColl = xmlUEle.attributes;
        xblur.textContent =
          attrColl.getNamedItem("w14:blurRad")?.value || xblur.textContent;
        xdist.textContent =
          attrColl.getNamedItem("w14:dist")?.value || xdist.textContent;
        xdir.textContent =
          attrColl.getNamedItem("w14:dir")?.value || xdir.textContent;
        xsx.textContent =
          attrColl.getNamedItem("w14:sx")?.value || xsx.textContent;
        xsy.textContent =
          attrColl.getNamedItem("w14:sy")?.value || xsy.textContent;

        const xmlSchemeColorNode = select(
          "w14:schemeClr",
          xmlUEle,
          true
        ) as Element;
        if (xmlSchemeColorNode) {
          xschemeclr.textContent = xmlSchemeColorNode.getAttribute("w14:val");

          let xmlTempNode = select(
            "w14:alpha",
            xmlSchemeColorNode,
            true
          ) as Element;
          if (xmlTempNode) {
            xlummod.textContent = xmlTempNode.getAttribute("w14:val");
          }

          xmlTempNode = select(
            "w14:lumMod",
            xmlSchemeColorNode,
            true
          ) as Element;
          if (xmlTempNode) {
            xlummod.textContent = xmlTempNode.getAttribute("w14:val");
          }
        }
      }
    } else {
      xblur.textContent = "0";
      xdist.textContent = "0";
      xdir.textContent = "0";
      xsx.textContent = "0";
      xsy.textContent = "0";
      xschemeclr.textContent = "none";
      xlummod.textContent = "0";
      xlumoff.textContent = "0";
    }
  };

  protected fillRowPropsStructureFromBaseFirst(
    strBasedOn: string,
    paramark: Element
  ): void {
    const strXPath = `/w:styles/w:style[@w:type='character' and @w:styleId='${strBasedOn}']`;

    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });

    if (!this.m_myPackageReader) {
      throw new Error("Package not initialized !!");
    }

    const m_StylesDoc = this.m_myPackageReader.partFileMap.get(
      "m_StylesDoc"
    ) as XMLDocument;
    const xmlCStyleNode = select(strXPath, m_StylesDoc, true) as Element;

    if (!xmlCStyleNode) {
      return;
    }

    const strBasedOnNew = WordUtils.extractStringFromVal(
      xmlCStyleNode,
      "w:basedOn",
      select
    );

    if (strBasedOnNew !== "") {
      this.fillRowPropsStructureFromBaseFirst(strBasedOnNew, paramark);
    }

    const xmlrPrEle = select("w:rPr", xmlCStyleNode, true) as Element;

    if (xmlrPrEle) {
      this.extractProperties(xmlrPrEle, paramark);
    }
  }

  public extractProperties = (
    xmlRPropsEle: Element,
    paramark: Element
  ): void => {
    const xSz = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "sz"
    );
    const xB = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "b"
    );
    const xI = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "i"
    );
    const xSmallCaps = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "smallcaps"
    );
    const xCaps = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "caps"
    );
    const xStrike = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "strike"
    );
    const xOutline = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "outline"
    );
    const xDStrike = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "dstrike"
    );
    const xShadow = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "shadow"
    );
    const xHidden = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "hidden"
    );
    const xEngrave = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "engrave"
    );
    const xEmboss = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "emboss"
    );
    const xVAlign = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "valign"
    );
    const xStyle = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "style"
    );
    const xSpac = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "spac"
    );
    const xWd = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "wd"
    );
    const xKern = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "kern"
    );
    const xPos = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "pos"
    );

    this.generateFontNameNode(xmlRPropsEle, paramark);

    const select = xpath.useNamespaces({
      w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    });

    xSz.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:sz", select) === ""
        ? xSz.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:sz", select);

    // val attribute may be present if defaults are applied (default style)
    xB.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:b", select) === ""
        ? xB.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:b", select);
    xB.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:b", select) === ""
        ? xB.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:b", select);

    xI.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:i", select) === ""
        ? xI.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:i", select);
    xI.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:i", select) === ""
        ? xI.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:i", select);

    xSmallCaps.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:smallCaps", select) === ""
        ? xSmallCaps.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:smallCaps", select);
    xSmallCaps.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:smallCaps", select) === ""
        ? xSmallCaps.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:smallCaps", select);

    xCaps.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:caps", select) === ""
        ? xCaps.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:caps", select);
    xCaps.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:caps", select) === ""
        ? xCaps.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:caps", select);

    xStrike.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:strike", select) === ""
        ? xStrike.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:strike", select);
    xStrike.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:strike", select) === ""
        ? xStrike.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:strike", select);

    xOutline.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:outline", select) === ""
        ? xOutline.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:outline", select);
    xOutline.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:outline", select) === ""
        ? xOutline.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:outline", select);

    xDStrike.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:dstrike", select) === ""
        ? xDStrike.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:dstrike", select);
    xDStrike.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:dstrike", select) === ""
        ? xDStrike.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:dstrike", select);

    xHidden.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:vanish", select) === ""
        ? xHidden.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:vanish", select);
    xHidden.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:vanish", select) === ""
        ? xHidden.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:vanish", select);

    xEngrave.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:imprint", select) === ""
        ? xEngrave.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:imprint", select);
    xEngrave.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:imprint", select) === ""
        ? xEngrave.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:imprint", select);

    xEmboss.textContent =
      WordUtils.extractString(xmlRPropsEle, "w:emboss", select) === ""
        ? xEmboss.textContent
        : WordUtils.extractString(xmlRPropsEle, "w:emboss", select);
    xEmboss.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:emboss", select) === ""
        ? xEmboss.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:emboss", select);

    xVAlign.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:vertAlign", select) === ""
        ? xVAlign.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:vertAlign", select);
    xStyle.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:rStyle", select) === ""
        ? xStyle.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:rStyle", select);
    xSpac.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:spacing", select) === ""
        ? xSpac.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:spacing", select);
    xWd.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:w", select) === ""
        ? xWd.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:w", select);
    xKern.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:kern", select) === ""
        ? xKern.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:kern", select);
    xPos.textContent =
      WordUtils.extractStringFromVal(xmlRPropsEle, "w:position", select) === ""
        ? xPos.textContent
        : WordUtils.extractStringFromVal(xmlRPropsEle, "w:position", select);

    //Paramark Underline Node
    this.generateUndelineNode(xmlRPropsEle, paramark);

    //Paramark Color Node
    this.generateColorNode(xmlRPropsEle, paramark);

    //Paramark Shadow node
    this.generateShadowNode(xmlRPropsEle, paramark);

    //assign Default values
    xSz.textContent = xSz.textContent == "" ? "20" : xSz.textContent;

    xB.textContent = xB.textContent == "" ? "0" : xB.textContent;

    xI.textContent = xI.textContent == "" ? "0" : xI.textContent;

    xSmallCaps.textContent =
      xSmallCaps.textContent == "" ? "0" : xSmallCaps.textContent;

    xCaps.textContent = xCaps.textContent == "" ? "0" : xCaps.textContent;

    xStrike.textContent = xStrike.textContent == "" ? "0" : xStrike.textContent;

    xOutline.textContent =
      xOutline.textContent == "" ? "0" : xOutline.textContent;

    xDStrike.textContent =
      xDStrike.textContent == "" ? "0" : xDStrike.textContent;

    let xTemp = paramark.getElementsByTagName("shadow/blur")[0];

    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "0" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("shadow/dist")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "0" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("shadow/angle")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "0" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("shadow/sizex")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "0" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("shadow/sizey")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "0" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("shadow/schemeclr")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "0" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("shadow/lummod")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "0" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("shadow/lumoff")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;
    }

    xHidden.textContent = xHidden.textContent == "" ? "0" : xHidden.textContent;

    xEngrave.textContent =
      xEngrave.textContent == "" ? "0" : xEngrave.textContent;

    xEmboss.textContent = xEmboss.textContent == "" ? "0" : xEmboss.textContent;

    xVAlign.textContent =
      xVAlign.textContent == "" ? "none" : xVAlign.textContent;

    xStyle.textContent = xStyle.textContent == "" ? "none" : xStyle.textContent;

    xSpac.textContent = xSpac.textContent == "" ? "0" : xSpac.textContent;

    xWd.textContent = xWd.textContent == "" ? "0" : xWd.textContent;

    xKern.textContent = xKern.textContent == "" ? "0" : xKern.textContent;

    xPos.textContent = xPos.textContent == "" ? "0" : xPos.textContent;

    xTemp = paramark.getElementsByTagName("u/value")[0];
    xTemp.textContent = xTemp.textContent === "" ? "0" : xTemp.textContent;

    xTemp = paramark.getElementsByTagName("u/type")[0];
    xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;

    xTemp = paramark.getElementsByTagName("u/clr")[0];
    xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;

    xTemp = paramark.getElementsByTagName("u/themeclr")[0];
    xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;

    xTemp = paramark.getElementsByTagName("u/themeshade")[0];
    xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;

    xTemp = paramark.getElementsByTagName("u/themetint")[0];
    xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;

    xTemp = paramark.getElementsByTagName("clr/hex")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("clr/themeclr")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("clr/themeshade")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;
    }

    xTemp = paramark.getElementsByTagName("clr/themetint")[0];
    if (xTemp) {
      xTemp.textContent = xTemp.textContent === "" ? "none" : xTemp.textContent;
    }
  };

  protected generateFontNameNode = (
    office12XMLNode: Element,
    paramark: Element
  ): void => {
    let strXPath = "",
      xElement;
    const select = xpath.useNamespaces({
      a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    });
    if (!this.m_myPackageReader) {
      throw new Error("Package not initialized !!");
    }
    const xmlDoc = this.m_myPackageReader.partFileMap.get("m_ThemeDoc");

    if (office12XMLNode) {
      const xmlFontEle = select("w:rFonts", office12XMLNode, true) as Element;

      if (xmlFontEle) {
        const attrColl = xmlFontEle.attributes;

        if (attrColl.getNamedItem("w:hAnsi")) {
          WordUtils.returnNode(
            this.m_myXmlDoc as Document,
            paramark,
            "name"
          ).textContent = attrColl.getNamedItem("w:hAnsi")?.value || "";
        } else {
          let fontNode = attrColl.getNamedItem("w:hAnsiTheme");
          if (fontNode !== null) {
            let sFontTest = fontNode.textContent; // In TypeScript, innerText is typically textContent
            switch (sFontTest) {
              case "minorHAnsi": {
                let strXPath =
                  "/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin";

                const select = xpath.useNamespaces({
                  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
                });
                let xElement = select(
                  strXPath,
                  this.m_myPackageReader.partFileMap.get("m_ThemeDoc") as Node,
                  true
                ) as Element;

                if (xElement !== null) {
                  let nodeName = WordUtils.returnNode(
                    this.m_myXmlDoc as XMLDocument,
                    paramark,
                    "name"
                  );
                  nodeName.textContent =
                    xElement.getAttribute("typeface") || "";
                }

                break;
              }
              case "majorHAnsi": {
                let strXPath =
                  "/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin";
                const select = xpath.useNamespaces({
                  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
                });
                let xElement = select(
                  strXPath,
                  this.m_myPackageReader.partFileMap.get("m_ThemeDoc") as Node
                ) as Element;
                select;
                if (xElement !== null) {
                  let nodeName = WordUtils.returnNode(
                    this.m_myXmlDoc as XMLDocument,
                    paramark,
                    "name"
                  );
                  nodeName.textContent =
                    xElement.getAttribute("typeface") || "";
                }

                break;
              }
            }
          }
        }
      } else {
        //resolving sonar issue
        WordUtils.returnNode(this.m_myXmlDoc as XMLDocument, paramark, "name");
      }
    } else {
      //resolving sonar issue
      WordUtils.returnNode(this.m_myXmlDoc as XMLDocument, paramark, "name");
    }

    strXPath = "/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin";

    // Retrieve the 'm_ThemeDoc' from the package's file map
    const m_ThemeDoc = this.m_myPackageReader.partFileMap.get(
      "m_ThemeDoc"
    ) as XMLDocument;
    let strThemeMinorFont: string | null = null;

    // Execute the XPath query and get the 'typeface' attribute
    xElement = select(strXPath, m_ThemeDoc, true) as Element;
    strThemeMinorFont = xElement.getAttribute("typeface");

    // Update the inner text of the 'name' node if it is empty
    const nameNode = paramark.getElementsByTagName("name")[0];
    if (nameNode) {
      nameNode.textContent =
        nameNode.textContent === "" ? strThemeMinorFont : nameNode.textContent;
    }
  };

  protected generateUndelineNode = (
    xmlRPropsEle: Element,
    paramark: Element
  ): void => {
    const underline = WordUtils.returnNode(
      this.m_myXmlDoc as XMLDocument,
      paramark,
      "u"
    );
    let xmlUEle = null;
    if (xmlRPropsEle != null) {
      const select = xpath.useNamespaces({
        w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      });

      xmlUEle = select("w:u", xmlRPropsEle, true) as Element;
      if (xmlUEle != null) {
        const attrColl = xmlUEle.attributes;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "value"
        ).textContent =
          attrColl.getNamedItem("w:val") != null
            ? "1"
            : WordUtils.returnNode(
                this.m_myXmlDoc as XMLDocument,
                underline,
                "value"
              ).textContent;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "type"
        ).textContent =
          attrColl?.getNamedItem("w:val")?.value ||
          WordUtils.returnNode(
            this.m_myXmlDoc as XMLDocument,
            underline,
            "type"
          ).textContent;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "clr"
        ).textContent = attrColl.getNamedItem("w:color")?.value || "Automatic";
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "themeclr"
        ).textContent =
          attrColl.getNamedItem("w:themeColor")?.value ||
          WordUtils.returnNode(
            this.m_myXmlDoc as XMLDocument,
            underline,
            "themeclr"
          ).textContent;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "themeshade"
        ).textContent =
          attrColl.getNamedItem("w:themeShade")?.value ||
          WordUtils.returnNode(
            this.m_myXmlDoc as XMLDocument,
            underline,
            "themeshade"
          ).textContent;
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "themetint"
        ).textContent =
          attrColl.getNamedItem("w:themeTint")?.value ||
          WordUtils.returnNode(
            this.m_myXmlDoc as XMLDocument,
            underline,
            "themetint"
          ).textContent;
      } else {
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "value"
        );
        WordUtils.returnNode(this.m_myXmlDoc as XMLDocument, underline, "type");
        WordUtils.returnNode(this.m_myXmlDoc as XMLDocument, underline, "clr");
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "themeclr"
        );
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "themeshade"
        );
        WordUtils.returnNode(
          this.m_myXmlDoc as XMLDocument,
          underline,
          "themetint"
        );
      }
    } else {
      WordUtils.returnNode(this.m_myXmlDoc as XMLDocument, underline, "value");
      WordUtils.returnNode(this.m_myXmlDoc as XMLDocument, underline, "type");
      WordUtils.returnNode(this.m_myXmlDoc as XMLDocument, underline, "clr");
      WordUtils.returnNode(
        this.m_myXmlDoc as XMLDocument,
        underline,
        "themeclr"
      );
      WordUtils.returnNode(
        this.m_myXmlDoc as XMLDocument,
        underline,
        "themeshade"
      );
      WordUtils.returnNode(
        this.m_myXmlDoc as XMLDocument,
        underline,
        "themetint"
      );
    }
  };
}
