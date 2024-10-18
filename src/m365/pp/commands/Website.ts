import GlobalOptions from "../../../GlobalOptions";
export interface PpWebSiteOptions extends GlobalOptions {
  environmentName: string;
  id?: string;
  name?: string;
  url?: string;
  asAdmin?: boolean;
}