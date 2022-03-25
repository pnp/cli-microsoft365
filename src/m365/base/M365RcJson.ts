export interface M365RcJson {
  apps?: M365RcJsonApp[];
}

export interface M365RcJsonApp {
  appId: string;
  name: string;
}