import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  VivaConnectionsDefaultStart?: boolean;
}

class SpoHomeSiteSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_SET;
  }

  public get description(): string {
    return 'Sets the specified site as the Home Site and optionally set the Viva Connections landing experience to the SharePoint Home Site';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-v, --VivaConnectionsDefaultStart [VivaConnectionsDefaultStart]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);
      let requestData: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">`;
      if (args.options.VivaConnectionsDefaultStart) {
        requestData += `<Actions><Method Name="ValidateMultipleHomeSitesParameterExists" Id="85" ObjectPathId="81"><Parameters><Parameter Type="Boolean">false</Parameter></Parameters></Method><Method Name="ValidateVivaHomeParameterExists" Id="86" ObjectPathId="81"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method><Method Name="SetSPHSiteWithConfigurations" Id="87" ObjectPathId="81"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.siteUrl)}</Parameter><Parameter Type="Boolean">${args.options.VivaConnectionsDefaultStart}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="81" Name="b6e793a0-e066-6000-3c4a-cb1f897402b4|908bed80-a04a-4433-b4a0-883d9847d110:d872ec63-6bea-4678-9429-078f4fa93560&#xA;Tenant" /></ObjectPaths></Request>`;
      }
      else {
        requestData += `<Actions><ObjectPath Id="57" ObjectPathId="56" /><Method Name="SetSPHSite" Id="58" ObjectPathId="56"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.siteUrl)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="56" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;

      }
      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: requestData
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        logger.log(json[json.length - 1]);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoHomeSiteSetCommand();