declare interface IRouteCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'RouteCommandSetStrings' {
  const strings: IRouteCommandSetStrings;
  export = strings;
}
