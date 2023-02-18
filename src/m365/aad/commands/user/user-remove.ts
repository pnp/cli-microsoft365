import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { Cli } from '../../../../cli/Cli';
import GraphCommand from '../../../base/GraphCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  userName?: string;
  confirm?: boolean;
}

class AadUserRemoveCommand extends GraphCommand {

  public get name(): string {
    return commands.USER_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific user';
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
        name: typeof args.options.name !== 'undefined',
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id [id]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'userName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid user principal name (UPN)`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing user '${args.options.id || args.options.userName}'...`);
    }

    if (args.options.confirm) {
      await this.deleteUser(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove user '${args.options.id || args.options.userName}'?`
      });

      if (result.continue) {
        await this.deleteUser(args);
      }
    }
  }

  private async deleteUser(args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users/${args.options.id || args.options.userName}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadUserRemoveCommand();