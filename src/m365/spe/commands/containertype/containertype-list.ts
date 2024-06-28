import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ContainerTypeProperties } from '../../ContainerTypeProperties.js';

class SpeContainertypeListCommand extends SpoCommand {
  private allContainerTypes?: ContainerTypeProperties[];

  public get name(): string {
    return commands.CONTAINERTYPE_LIST;
  }

  public get description(): string {
    return 'Lists all Container Types';
  }

  public defaultProperties(): string[] | undefined {
    return ['ContainerTypeId', 'DisplayName', 'OwningAppId'];
  }
  
  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving list of Container types...`);
      }

      this.allContainerTypes = [];

      await this.getAllContainerTypes(spoAdminUrl, logger);

      await logger.log(this.allContainerTypes);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getAllContainerTypes(spoAdminUrl: string, logger: Logger): Promise<void> {
    const formDigest: FormDigestInfo | undefined = undefined;
    const formDigestInfo: FormDigestInfo = await spo.ensureFormDigest(spoAdminUrl, logger, formDigest, this.debug);

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestInfo.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="46" ObjectPathId="45" /><Method Name="GetSPOContainerTypes" Id="47" ObjectPathId="45"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="45" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const res: string = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      throw (response.ErrorInfo.ErrorMessage);
    }
    else {
      const containerTypes: ContainerTypeProperties[] = json[json.length - 1];
      this.allContainerTypes!.push(...containerTypes);
    }
  }
}

export default new SpeContainertypeListCommand();