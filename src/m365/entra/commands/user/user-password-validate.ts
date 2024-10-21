import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  password: string;
}

class EntraUserPasswordValidateCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_PASSWORD_VALIDATE;
  }

  public get description(): string {
    return "Check a user's password against the organization's password validation policy";
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --password <password>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/users/validatePassword`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: {
          password: args.options.password
        },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserPasswordValidateCommand();