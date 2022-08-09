import { Cli, Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, spo } from '../../../../utils';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const cdnTypeString: string = args.options.type || 'Public';
    const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;
    let spoAdminUrl: string = '';
    let tenantId: string = '';

    const removeCdnOrigin = (): void => {
      spo
        .getTenantId(logger, this.debug)
        .then((_tenantId: string): Promise<string> => {
          tenantId = _tenantId;
          return spo.getSpoAdminUrl(logger, this.debug);
        })
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return spo.getRequestDigest(spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            logger.logToStderr(`Removing origin ${args.options.origin} from the ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
          }

          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="RemoveTenantCdnOrigin" Id="33" ObjectPathId="29"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.origin)}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="29" Name="${tenantId}" /></ObjectPaths></Request>`
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            cb(new CommandError(response.ErrorInfo.ErrorMessage));
          }
          else {
            cb();
          }
        }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      if (this.debug) {
        logger.logToStderr('Confirmation suppressed through the confirm option. Removing CDN origin...');
      }
      removeCdnOrigin();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to delete the ${args.options.origin} CDN origin?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeCdnOrigin();
        }
      });
    }
  }
}

module.exports = new SpoCdnOriginRemoveCommand();