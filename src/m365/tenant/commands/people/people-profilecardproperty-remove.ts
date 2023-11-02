import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { profileCardPropertyNames } from './profileCardProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  force?: boolean;
}

class TenantPeopleProfileCardPropertyRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.PEOPLE_PROFILECARDPROPERTY_REMOVE;
  }

  public get description(): string {
    return 'Removes a custom attribute as a profile card property';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: args.options.name
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>',
        autocomplete: profileCardPropertyNames
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (profileCardPropertyNames.every(n => n.toLowerCase() !== args.options.name.toLowerCase())) {
          return `${args.options.name} is not a valid value for name. Allowed values are ${profileCardPropertyNames.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeProfileCardProperty = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing '${args.options.name}' as a profile card property...`);
      }

      const requestOptions: any = {
        url: `${this.resource}/v1.0/admin/people/profileCardProperties/${args.options.name}`,
        headers: {
          'content-type': 'application/json'
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
      await removeProfileCardProperty();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the profile card property '${args.options.name}'?`
      });
      if (result.continue) {
        await removeProfileCardProperty();
      }
    }
  }
}

export default new TenantPeopleProfileCardPropertyRemoveCommand();