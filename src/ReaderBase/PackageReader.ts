import JSZip from "jszip";
import xmldom from "@xmldom/xmldom";

interface relationship {
  id: string | null;
  type: string | null;
  target: string | null;
  targetMode: string | null;
}

export class PackageReader {
  private package: JSZip | null = null;
  public partFileMap: Map<string, XMLDocument>; //need to check type definition.

  constructor() {
    this.partFileMap = new Map();
  }

  initalisePackage = async (
    strFileName: Buffer,
    checkForIncompatibleOfficePlatform: boolean
  ): Promise<boolean> => {
    if (strFileName != null) {
      try {
        // MemoryStream strm = new MemoryStream(); - need alternative
        // CopyStream(strFileName, strm);
        // strFileName.Close();
        var zip = new JSZip();
        this.package = await zip.loadAsync(strFileName);
        if (checkForIncompatibleOfficePlatform) {
          // CheckForIncompatibleOfficeApp(package);
        }
        return true;
      } catch (err) {
        throw err;
      }
    } else {
      return false;
    }
  };

  ReturnBaseXML = async (uri: string | null, schemaName: string) => {
    if (!this.package) {
      throw new Error("Package is not initailized yet..");
    }

    let relColl: Array<relationship> | null = null;

    if (!uri) {
      relColl = await this.getRelationships();
    } else {
      const part = this.package.file(uri.replace(/^\//, "")); // PackagePart part = package.GetPart(uri);
      if (part == null) {
        throw new Error("Part is not present in the package");
      }
      relColl = await this.getRelationships("word", "document.xml");
    }

    return this.findTargetURI(relColl, schemaName);
  };

  ReturnPackagePart = async (
    uri: string,
    lenLimit: number = 0
  ): Promise<Buffer | null> => {
    if (!this.package) {
      throw new Error("Package is not initailized yet..");
    }
    const file = this.package.file(uri.replace(/^\//, ""));

    if (!file) return null;

    const contentBuffer = await file.async("nodebuffer");
    if (lenLimit === 0 || contentBuffer.length <= lenLimit) {
      return contentBuffer;
    } else {
      throw new Error(
        "Failed to parse the submitted document size greator than requested size. Check the document."
      );
    }
  };

  ReturnPackageProperties = () => {
    if (!this.package) {
      throw new Error("Package is not initailized yet..");
    }
    return this.package;
  };

  ReturnDocmentFromPart = async (
    sFilePartName: string,
    lenlimit: number = 0
  ) => {
    if (!this.package) {
      throw new Error("Package is not initailized yet..");
    }

    if (sFilePartName == "") {
      throw new Error("Part file name is not specified");
    }

    const sFileStream: Buffer | null = await this.ReturnPackagePart(
      sFilePartName,
      lenlimit
    );

    let doc: XMLDocument = new xmldom.DOMImplementation().createDocument(
      null,
      null,
      null
    );

    if (sFileStream) {
      const parser = new xmldom.DOMParser();
      doc = parser.parseFromString(sFileStream.toString(), "text/xml");
    }

    return doc;
  };

  private findTargetURI(
    relationships: relationship[] | null,
    schemaName: string
  ): string {
    if (!relationships?.length) return "";
    for (let i = 0; i < relationships.length; i++) {
      if (relationships[i].type === schemaName) {
        return relationships[i].target || "";
      }
    }
    return "";
  }

  private getRelationships = async (
    partUriFolder: string | null = null,
    partUriXML: string | null = null
  ): Promise<Array<relationship>> => {
    if (!this.package) {
      throw new Error("Package is not initailized yet..");
    }
    const parser = new xmldom.DOMParser();
    let relationships = [];

    // Determine the path to the relationships file
    const relsPath = partUriFolder
      ? `${partUriFolder}/_rels/${partUriXML}.rels`
      : "_rels/.rels";

    const relsFile = this.package.file(relsPath);
    if (!relsFile) {
      console.log("No relationships file found at the specified path.");
      return []; // Return an empty array if no relationships file is found
    }

    // Read and parse the relationships XML
    const xmlData = await relsFile.async("string");
    const doc = parser.parseFromString(xmlData, "application/xml");
    const relationshipElements = doc.getElementsByTagName("Relationship");

    for (let i = 0; i < relationshipElements.length; i++) {
      const rel = relationshipElements[i];
      relationships.push({
        id: rel.getAttribute("Id"),
        type: rel.getAttribute("Type"),
        target: rel.getAttribute("Target"),
        targetMode: rel.getAttribute("TargetMode"),
      });
    }

    return relationships;
  };
}
