import { Cli, Logger } from '../../../../cli';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, formatting, spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type: string;
  origin: string;
  confirm?: boolean;
}

class SpoCdnOriginRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.CDN_ORIGIN_REMOVE;
  }

  public get description(): string {
    return 'Removes CDN origin for the current SharePoint Online tenant';
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
        cdnType: args.options.type || 'Public',
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type [type]',
        autocomplete: ['Public', 'Private']
      },
      {
        option: '-r, --origin <origin>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type) {
          if (args.options.type !== 'Public' &&
            args.options.type !== 'Private') {
            return `${args.options.type} is not a valid CDN type. Allowed values are Public|Private`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const cdnTypeString: string = args.options.type || 'Public';
    const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;

    const removeCdnOrigin: () => Promise<void> = async (): Promise<void> => {
      try {
        const tenantId = await spo.getTenantId(logger, this.debug);
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const reqDigest = await spo.getRequestDigest(spoAdminUrl);

        if (this.verbose) {
          logger.logToStderr(`Removing origin ${args.options.origin} from the ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.origin)}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="${tenantId}" /></ObjectPaths></Request>`
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

    if (args.options.confirm) {
      if (this.debug) {
        logger.logToStderr('Confirmation suppressed through the confirm option. Removing CDN origin...');
      }
      await removeCdnOrigin();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to delete the ${args.options.origin} CDN origin?`
      });

      if (result.continue) {
        await removeCdnOrigin();
      }
    }
  }
}

module.exports = new SpoCdnOriginRemoveCommand();