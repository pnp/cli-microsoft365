export interface Group {
  id: string;
  deletedDateTime?: string;
  classification?: string;
  createdDateTime: string;
  description?: string;
  displayName: string;
  groupTypes: string[],
  mail: string;
  mailEnabled: boolean;
  mailNickname: string;
  onPremisesLastSyncDateTime?: string;
  onPremisesProvisioningErrors: string[];
  onPremisesSecurityIdentifier?: string;
  onPremisesSyncEnabled?: boolean;
  owners: any;
  preferredDataLocation?: string;
  proxyAddresses: string[];
  renewedDateTime: string;
  securityEnabled: boolean;
  siteUrl?: string;
  visibility: string;
}