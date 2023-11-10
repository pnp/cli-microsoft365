import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  userId: string;
  userName?: string;
  force?: boolean;
}

class TeamsUserAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_APP_REMOVE;
  }

  public get description(): string {
    return 'Uninstall an app from the personal scope of the specified user.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        force: (!!args.options.force).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id [id]'
      },
      {
        option: '--name [name]'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
    this.optionSets.push({ options: ['userId', 'userName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeApp = async (): Promise<void> => {
      const appId: string = await this.getAppId(args);
      // validation ensures that here we have either userId or userName
      const userId: string = (args.options.userId ?? args.options.userName) as string;
      const endpoint: string = `${this.resource}/v1.0`;

      if (this.verbose) {
        await logger.logToStderr(`Removing app with ID ${args.options.id} for user ${args.options.userId}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${endpoint}/users/${args.options.userId}/teamwork/installedApps/${appId}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeApp();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the app ${args.options.id || args.options.name} for user ${args.options.userId}?`
      });

      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to remove the app with id ${args.options.id} for user ${args.options.userId ?? args.options.userName}?` });
      if (result) {
        await removeApp();
      }
    }
  }

  private async getAppId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/users/${args.options.userId}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(args.options.name as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(requestOptions);
    const app: { id: string; } | undefined = response.value[0];

    if (!app) {
      throw `The specified Teams app does not exist`;
    }

    if (response.value.length > 1) {
      throw `Multiple Teams apps with name ${args.options.name} found. Please choose one of these ids: ${response.value.map(x => x.id).join(', ')}`;
    }

    return app.id;
  }
}

export default new TeamsUserAppRemoveCommand();