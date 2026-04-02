declare interface IAdobePdfWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ClientIdFieldLabel: string;
  ClientIdFieldDescription: string;
  FilePickerButtonLabel: string;
  FilePickerLabel: string;
  ViewModeFieldLabel: string;
  MissingClientId: string;
  MissingFile: string;
  SdkLoadError: string;
}

declare module 'AdobePdfWebPartStrings' {
  const strings: IAdobePdfWebPartStrings;
  export = strings;
}
