declare interface IRevisionHistoryWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  SiteUrlFieldLabel: string;
  listName: string;
  draft: "Draft";
  publishded: "Published";
  uV: "Under verification";
  verified: "Verified";
  published: "Published";
  cancel: "Cancelled";

}

declare module 'RevisionHistoryWebPartStrings' {
  const strings: IRevisionHistoryWebPartStrings;
  export = strings;
}
