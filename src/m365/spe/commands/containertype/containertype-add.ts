import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface ContainerType {
  _ObjectType_?: string;
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

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  name: string;
  applicationId: string;
  trial?: boolean;
  azureSubscriptionId?: string;
  resourceGroup?: string;
  region?: string;
}

class SpeContainerTypeAddCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTAINERTYPE_ADD;
  }

  public get description(): string {
    return 'Creates a new Containertype for your app';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        trial: !!args.options.trial,
        azureSubscriptionId: typeof args.options.azureSubscriptionId !== 'undefined',
        resourceGroup: typeof args.options.resourceGroup !== 'undefined',
        region: typeof args.options.region !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '--applicationId <applicationId>'
      },
      {
        option: '--trial'
      },
      {
        option: '--azureSubscriptionId [azureSubscriptionId]'
      },
      {
        option: '--resourceGroup [resourceGroup]'
      },
      {
        option: '--region [region]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.applicationId)) {
          return `${args.options.applicationId} is not a valid GUID`;
        }

        if (args.options.trial === undefined && !args.options.azureSubscriptionId) {
          return 'You must specify the azureSubscriptionId when creating a non-trial environment';
        }

        if (args.options.trial === undefined && !args.options.resourceGroup) {
          return 'You must specify the resourceGroup when creating a non-trial environment';
        }

        if (args.options.trial === undefined && !args.options.region) {
          return 'You must specify the region when creating a non-trial environment';
        }

        if (args.options.azureSubscriptionId && !validation.isValidGuid(args.options.azureSubscriptionId)) {
          return `${args.options.azureSubscriptionId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Creating a new Containertype for your app with name ${args.options.name}`);
      }

      const adminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const requestBody = this.getRequestBody(args.options);

      const formDigestInfo: FormDigestInfo = await spo.ensureFormDigest(adminUrl, logger, undefined, this.debug);
      const requestOptions: CliRequestOptions = {
        url: `${adminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': formDigestInfo.FormDigestValue
        },
        data: requestBody
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];

      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      const result: ContainerType = json.pop();

      delete result._ObjectType_;
      result.SPContainerTypeBillingClassification = args.options.trial ? 'Trial' : 'Standard';
      result.AzureSubscriptionId = this.replaceString(result.AzureSubscriptionId);
      result.OwningAppId = this.replaceString(result.OwningAppId);
      result.OwningTenantId = this.replaceString(result.OwningTenantId);
      result.ContainerTypeId = this.replaceString(result.ContainerTypeId);

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private replaceString(s: string): string {
    return s.replace('/Guid(', '').replace(')/', '');
  }

  private getRequestBody(options: Options): string {
    if (options.trial) {
      return `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="NewSPOContainerType" Id="5" ObjectPathId="3"><Parameters><Parameter TypeId="{5466648e-c306-441b-9df4-c09deef25cb1}"><Property Name="AzureSubscriptionId" Type="Guid">{00000000-0000-0000-0000-000000000000}</Property><Property Name="ContainerTypeId" Type="Guid">{00000000-0000-0000-0000-000000000000}</Property><Property Name="CreationDate" Type="Null" /><Property Name="DisplayName" Type="String">${options.name}</Property><Property Name="ExpiryDate" Type="Null" /><Property Name="IsBillingProfileRequired" Type="Boolean">false</Property><Property Name="OwningAppId" Type="Guid">{${options.applicationId}}</Property><Property Name="OwningTenantId" Type="Guid">{00000000-0000-0000-0000-000000000000}</Property><Property Name="Region" Type="Null" /><Property Name="ResourceGroup" Type="Null" /><Property Name="SPContainerTypeBillingClassification" Type="Enum">1</Property></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
    }
    return `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="NewSPOContainerType" Id="4" ObjectPathId="1"><Parameters><Parameter TypeId="{5466648e-c306-441b-9df4-c09deef25cb1}"><Property Name="AzureSubscriptionId" Type="Guid">{${options.azureSubscriptionId}}</Property><Property Name="ContainerTypeId" Type="Guid">{00000000-0000-0000-0000-000000000000}</Property><Property Name="CreationDate" Type="Null" /><Property Name="DisplayName" Type="String">${options.name}</Property><Property Name="ExpiryDate" Type="Null" /><Property Name="IsBillingProfileRequired" Type="Boolean">false</Property><Property Name="OwningAppId" Type="Guid">{${options.applicationId}}</Property><Property Name="OwningTenantId" Type="Guid">{00000000-0000-0000-0000-000000000000}</Property><Property Name="Region" Type="String">${options.region}</Property><Property Name="ResourceGroup" Type="String">${options.resourceGroup}</Property><Property Name="SPContainerTypeBillingClassification" Type="Enum">0</Property></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
  }
}

export default new SpeContainerTypeAddCommand();