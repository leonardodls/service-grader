export class CommonFunctions {
  static PrependStringToURIPath = (
    uripath: string,
    stringToPrepend: string
  ): string => {
    let newURIPath: string = uripath;

    if (!uripath.startsWith(stringToPrepend)) {
      newURIPath = stringToPrepend + uripath;
    }

    return newURIPath;
  };
}
