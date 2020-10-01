declare interface IExtensionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ExtensionCommandSetStrings' {
  const strings: IExtensionCommandSetStrings;
  export = strings;
}
