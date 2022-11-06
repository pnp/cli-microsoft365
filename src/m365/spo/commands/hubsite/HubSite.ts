import { AssociatedSite } from './AssociatedSite';

export interface HubSite {
  'odata.etag'?: string;
  AssociatedSites?: AssociatedSite[];
  Description: string;
  ID: string;
  LogoUrl: string;
  SiteId: string;
  SiteUrl: string;
  Targets: string;
  TenantInstanceId: string;
  Title: string;
}