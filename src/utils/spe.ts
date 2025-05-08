import request, { CliRequestOptions } from '../request.js';
import { ClientSvcResponse, ClientSvcResponseContents } from './spo.js';
import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';
import config from '../config.js';
import { odata } from './odata.js';

export interface ContainerTypeProperties {
  AzureSubscriptionId: string;
  ContainerTypeId: string;
  CreationDate: string;
  DisplayName: string;
  ExpiryDate: string;
  IsBillingProfileRequired: boolean;
  OwningAppId: string;
  OwningTenantId: string;
  Region?: string;
  ResourceGroup?: string;
  SPContainerTypeBillingClassification: string;
}

export interface ContainerProperties {
  id: string;
  displayName: string;
  containerTypeId: string;
  createdDateTime: string;
}

const graphResource = 'https://graph.microsoft.com';

export const spe = {
  /**
   * Get all container types.
   * @param spoAdminUrl The URL of the SharePoint Online admin center site (e.g. https://contoso-admin.sharepoint.com)
   * @returns Array of container types
   */
  async getAllContainerTypes(spoAdminUrl: string): Promise<ContainerTypeProperties[]> {
    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="46" ObjectPathId="45" /><Method Name="GetSPOContainerTypes" Id="47" ObjectPathId="45"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="45" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const json = await request.post<ClientSvcResponse>(requestOptions);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      throw new Error(response.ErrorInfo.ErrorMessage);
    }

    const containerTypes: ContainerTypeProperties[] = json[json.length - 1];
    // Format the response to remove CSOM GUIDs and convert them to real GUIDs
    containerTypes.forEach(ct => {
      delete (ct as any)._ObjectType_;
      ct.AzureSubscriptionId = formatting.extractCsomGuid(ct.AzureSubscriptionId);
      ct.ContainerTypeId = formatting.extractCsomGuid(ct.ContainerTypeId);
      ct.OwningAppId = formatting.extractCsomGuid(ct.OwningAppId);
      ct.OwningTenantId = formatting.extractCsomGuid(ct.OwningTenantId);
    });

    return containerTypes;
  },

  /**
   * Get the ID of a container type by its name.
   * @param spoAdminUrl SharePoint Online admin center URL (e.g. https://contoso-admin.sharepoint.com)
   * @param name Name of the container type to search for
   * @returns ID of the container type
   */
  async getContainerTypeIdByName(spoAdminUrl: string, name: string): Promise<string> {
    const allContainerTypes = await this.getAllContainerTypes(spoAdminUrl);
    const containerTypes = allContainerTypes.filter(ct => ct.DisplayName.toLowerCase() === name!.toLowerCase());

    if (containerTypes.length === 0) {
      throw new Error(`The specified container type '${name}' does not exist.`);
    }

    if (containerTypes.length > 1) {
      const containerTypeKeyValuePair = formatting.convertArrayToHashTable('ContainerTypeId', containerTypes);
      const containerType = await cli.handleMultipleResultsFound<ContainerTypeProperties>(`Multiple container types with name '${name}' found.`, containerTypeKeyValuePair);
      return containerType.ContainerTypeId;
    }

    return containerTypes[0].ContainerTypeId;
  },

  /**
   * Get the ID of a container by its name.
   * @param containerTypeId ID of the container type.
   * @param name Name of the container to search for.
   * @returns ID of the container.
   */
  async getContainerIdByName(containerTypeId: string, name: string): Promise<string> {
    const containers = await odata.getAllItems<ContainerProperties>(`${graphResource}/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`);
    const matchingContainers = containers.filter(c => c.displayName.toLowerCase() === name.toLowerCase());

    if (matchingContainers.length === 0) {
      throw new Error(`The specified container '${name}' does not exist.`);
    }

    if (matchingContainers.length > 1) {
      const containerKeyValuePair = formatting.convertArrayToHashTable('id', matchingContainers);
      const container = await cli.handleMultipleResultsFound<ContainerProperties>(`Multiple containers with name '${name}' found.`, containerKeyValuePair);
      return container.id;
    }

    return matchingContainers[0].id;
  }
};