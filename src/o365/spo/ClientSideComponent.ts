/**
 * Client side webpart object 
 * (retrieved via the _api/web/GetClientSideWebParts REST call)
 */
export interface ClientSideComponent {
  ComponentType: number;
  Id: string;
  Manifest: string;
  ManifestType: number;
  Name: string;
  Status: number;
}