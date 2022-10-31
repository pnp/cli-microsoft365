import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteDesignId: string;
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
        option: '-i, --siteDesignId <siteDesignId>'
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
        if (!validation.isValidGuid(args.options.siteDesignId)) {
          return `${args.options.siteDesignId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const revokePermissions: () => Promise<void> = async (): Promise<void> => {
      try {
        const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
        const requestDigest: ContextInfo = await spo.getRequestDigest(spoUrl);
        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`,
          headers: {
            'X-RequestDigest': requestDigest.FormDigestValue,
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            id: args.options.siteDesignId,
            principalNames: args.options.principals.split(',').map(p => p.trim())
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      } 
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await revokePermissions();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to revoke access to site design ${args.options.siteDesignId} from the specified users?`
      });
      
      if (result.continue) {
        await revokePermissions();
      }
    }
  }
}

module.exports = new SpoSiteDesignRightsRevokeCommand();
