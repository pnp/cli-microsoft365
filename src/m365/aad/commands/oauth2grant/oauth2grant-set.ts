import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption
} from '../../../../Command';
import AadCommand from '../../../base/AadCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  grantId: string;
  scope: string;
}

class AadOAuth2GrantSetCommand extends AadCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_SET;
  }

  public get description(): string {
    return 'Update OAuth2 permissions for the service principal';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Updating OAuth2 permissions...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/myorganization/oauth2PermissionGrants/${encodeURIComponent(args.options.grantId)}?api-version=1.6`,
      headers: {
        'content-type': 'application/json'
      },
      json: true,
      body: {
        "scope": args.options.scope
      }
    };

    request
      .patch(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --grantId <grantId>',
        description: 'objectId of OAuth2 permission grant to update'
      },
      {
        option: '-s, --scope <scope>',
        description: 'Permissions to grant'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new AadOAuth2GrantSetCommand();