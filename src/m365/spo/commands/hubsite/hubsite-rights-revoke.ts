import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  hubSiteUrl: string;
  principals: string;
  force?: boolean;
}

class SpoHubSiteRightsRevokeCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_RIGHTS_REVOKE;
  }

  public get description(): string {
    return 'Revokes rights to join sites to the specified hub site for one or more principals';
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
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --hubSiteUrl <hubSiteUrl>'
      },
      {
        option: '-p, --principals <principals>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.hubSiteUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const revokeRights = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Revoking rights for ${args.options.principals} from ${args.options.hubSiteUrl}...`);
        }

        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const reqDigest = await spo.getRequestDigest(spoAdminUrl);

        const principals: string = args.options.principals
          .split(',')
          .map(p => `<Object Type="String">${formatting.escapeXml(p.trim())}</Object>`)
          .join('');

        const requestOptions: CliRequestOptions = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="RevokeHubSiteRights" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.hubSiteUrl)}</Parameter><Parameter Type="Array">${principals}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };

        const res = await request.post<string>(requestOptions);

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.force) {
      await revokeRights();
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to revoke rights to join sites to the hub site ${args.options.hubSiteUrl} from the specified users?`);

      if (result) {
        await revokeRights();
      }
    }
  }
}

export default new SpoHubSiteRightsRevokeCommand();
