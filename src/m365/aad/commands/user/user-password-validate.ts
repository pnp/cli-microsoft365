import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  password: string;
}

class AadUserPasswordValidateCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_PASSWORD_VALIDATE;
  }

  public get description(): string {
    return "Check a user's password against the organization's password validation policy";
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/beta/users/validatePassword`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: {
        password: args.options.password
      },
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --password <password>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

}

module.exports = new AadUserPasswordValidateCommand();