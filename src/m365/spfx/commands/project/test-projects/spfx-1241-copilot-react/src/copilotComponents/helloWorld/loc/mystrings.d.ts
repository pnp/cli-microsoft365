declare interface IHelloWorldCopilotComponentStrings {
  ExpandButtonLabel: string;
  OpenSiteButtonLabel: string;
  FollowUpButtonLabel: string;
  ResizeButtonLabel: string;
  CompactButtonLabel: string;
  SiteBadgePrefix: string;
  ThemeBadgePrefix: string;
  ModeBadgePrefix: string;
  GreetingPrefix: string;
  UnknownTheme: string;
  DefaultDisplayMode: string;
  FollowUpMessage: string;
  WelcomeGreeting: string;
  DisplayModeLabel: string;
  ThemeLabel: string;
  MessageLabel: string;
  ExpandToFullscreenTitle: string;
}

declare module 'HelloWorldCopilotComponentStrings' {
  const strings: IHelloWorldCopilotComponentStrings;
  export = strings;
}
