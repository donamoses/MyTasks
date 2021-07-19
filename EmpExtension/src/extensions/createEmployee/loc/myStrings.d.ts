declare interface ICreateEmployeeCommandSetStrings {
  Command1: string;
  Command2: string;

}

declare module 'CreateEmployeeCommandSetStrings' {
  const strings: ICreateEmployeeCommandSetStrings;
  export = strings;
}
