import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId: string;
  userId: string;
}

class TeamsUserAppAddCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_APP_ADD;
  }

  public get description(): string {
    return 'Install an app in the personal scope of the specified user';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--appId <appId>'
      },
      {
        option: '--userId <userId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        if (!validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0`;

    const requestOptions: any = {
      url: `${endpoint}/users/${args.options.userId}/teamwork/installedApps`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        'teamsApp@odata.bind': `${endpoint}/appCatalogs/teamsApps/${args.options.appId}`
      }
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (res: any): void => this.handleRejectedODataJsonPromise(res, logger, cb));
  }
}

module.exports = new TeamsUserAppAddCommand();