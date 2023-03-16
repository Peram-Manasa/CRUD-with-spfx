declare interface ISharepointframeWebPartStrings {
  ListFieldLabel: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  webURLFieldLabel:string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
}

declare module 'SharepointframeWebPartStrings' {
  const strings: ISharepointframeWebPartStrings;
  export = strings;
}
