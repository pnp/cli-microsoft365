import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class AadO365GroupRenewCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_RENEW;
  }

  public get description(): string {
    return `Renews Microsoft 365 group's expiration`;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Renewing Microsoft 365 group's expiration: ${args.options.id}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups/${args.options.id}/renew/`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
    };

    request
      .post(requestOptions)
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
        option: '-i, --id <id>',
        description: 'The ID of the Microsoft 365 group to renew'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new AadO365GroupRenewCommand();