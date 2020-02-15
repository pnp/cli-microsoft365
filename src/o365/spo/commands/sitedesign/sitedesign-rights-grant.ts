import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    return `${commands.SITEDESIGN_RIGHTS_GRANT}`;
  }

  public get description(): string {
    return 'Grants access to a site design for one or more principals';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let spoUrl: string = '';

    this
      .getSpoUrl(cmd, this.debug)
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
          body: {
            id: args.options.id,
            principalNames: args.options.principals.split(',').map(p => p.trim()),
            grantedRights: grantedRights
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the site design to grant rights on'
      },
      {
        option: '-p, --principals <principals>',
        description: 'Comma-separated list of principals to grant view rights. Principals can be users or mail-enabled security groups in the form of "alias" or "alias@<domain name>.com"'
      },
      {
        option: '-r, --rights <rights>',
        description: 'Rights to grant to principals. Available values View',
        autocomplete: ['View']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (!args.options.principals) {
        return 'Required parameter principals missing';
      }

      if (!args.options.rights) {
        return 'Required parameter rights missing';
      }

      if (args.options.rights !== 'View') {
        return `${args.options.rights} is not a valid rights value. Allowed values View`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Grant user with alias ${chalk.grey('PattiF')} view permission to the specified site design
      ${this.name} --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --principals PattiF --rights View

    Grant users with aliases ${chalk.grey('PattiF')} and ${chalk.grey('AdeleV')} view permission to the specified site design
      ${this.name} --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --principals "PattiF,AdeleV" --rights View

    Grant user with email ${chalk.grey('PattiF@contoso.com')} view permission to the specified site design
      ${this.name} --id 9b142c22-037f-4a7f-9017-e9d8c0e34b98 --principals PattiF@contoso.com --rights View

  More information:

    SharePoint site design and site script overview
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview
`);
  }
}

module.exports = new SpoSiteDesignRightsGrantCommand();
