import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  principals: string;
  confirm?: boolean;
}

class SpoSiteDesignRightsRevokeCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_RIGHTS_REVOKE;
  }

  public get description(): string {
    return 'Revokes access from a site design for one or more principals';
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
        confirm: args.options.confirm || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-p, --principals <principals>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const revokePermissions: () => void = (): void => {
      let spoUrl: string = '';

      spo
        .getSpoUrl(logger, this.debug)
        .then((_spoUrl: string): Promise<ContextInfo> => {
          spoUrl = _spoUrl;
          return spo.getRequestDigest(spoUrl);
        })
        .then((res: ContextInfo): Promise<void> => {
          const requestOptions: any = {
            url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`,
            headers: {
              'X-RequestDigest': res.FormDigestValue,
              'content-type': 'application/json;charset=utf-8',
              accept: 'application/json;odata=nometadata'
            },
            data: {
              id: args.options.id,
              principalNames: args.options.principals.split(',').map(p => p.trim())
            },
            responseType: 'json'
          };

          return request.post(requestOptions);
        })
        .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      revokePermissions();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to revoke access to site design ${args.options.id} from the specified users?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          revokePermissions();
        }
      });
    }
  }
}

module.exports = new SpoSiteDesignRightsRevokeCommand();
