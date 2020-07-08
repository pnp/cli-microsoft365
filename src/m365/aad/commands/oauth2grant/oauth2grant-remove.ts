import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import AadCommand from '../../../base/AadCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  grantId: string;
}

class AadOAuth2GrantRemoveCommand extends AadCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_REMOVE;
  }

  public get description(): string {
    return 'Remove specified service principal OAuth2 permissions';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Removing OAuth2 permissions...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/myorganization/oauth2PermissionGrants/${encodeURIComponent(args.options.grantId)}?api-version=1.6`,
      json: true
    };

    request
      .delete(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --grantId <grantId>',
        description: 'objectId of OAuth2 permission grant to remove'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.grantId) {
        return 'Required option grantId missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.OAUTH2GRANT_REMOVE).helpInformation());
    log(
      `  Remarks:
  
    Before you can remove service principal's OAuth2 permissions, you need to
    get the ${chalk.grey('objectId')} of the permissions grant to remove. You can retrieve it
    using the ${chalk.blue(commands.OAUTH2GRANT_LIST)} command.

    If the ${chalk.grey('objectId')} listed when using the ${chalk.blue(commands.OAUTH2GRANT_LIST)} command has a 
    minus sign ('-') prefix, you may receive an error indicating --grantId is
    missing. To resolve this issue simply  escape the leading '-',
    eg. ${chalk.blue(commands.OAUTH2GRANT_REMOVE)} --grantId \\-Zc1JRY8REeLxmXz5KtixAYU3Q6noCBPlhwGiX7pxmU     

  Examples:
  
    Remove the OAuth2 permission grant with ID ${chalk.grey('YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek')}
      ${commands.OAUTH2GRANT_REMOVE} --grantId YgA60KYa4UOPSdc-lpxYEnQkr8KVLDpCsOXkiV8i-ek

  More information:
  
    Application and service principal objects in Azure Active Directory (Azure AD)
      https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects
`);
  }
}

module.exports = new AadOAuth2GrantRemoveCommand();