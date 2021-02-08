import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ContextInfo } from '../../spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  principals: string;
  rights: string;
}

class SpoSiteDesignRightsGrantCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_RIGHTS_GRANT;
  }

  public get description(): string {
    return 'Grants access to a site design for one or more principals';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let spoUrl: string = '';

    this
      .getSpoUrl(logger, this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;
        return this.getRequestDigest(spoUrl);
      })
      .then((res: ContextInfo): Promise<void> => {
        const grantedRights: string = '1';
        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            id: args.options.id,
            principalNames: args.options.principals.split(',').map(p => p.trim()),
            grantedRights: grantedRights
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>'
      },
      {
        option: '-p, --principals <principals>'
      },
      {
        option: '-r, --rights <rights>',
        autocomplete: ['View']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.rights !== 'View') {
      return `${args.options.rights} is not a valid rights value. Allowed values View`;
    }

    return true;
  }
}

module.exports = new SpoSiteDesignRightsGrantCommand();
