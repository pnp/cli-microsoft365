export interface M365RcJson {
  apps?: M365RcJsonApp[];
  context?: any;
}

export interface M365RcJsonApp {
  appId: string;
  name: string;
}