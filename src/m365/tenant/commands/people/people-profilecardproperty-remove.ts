import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
    return 'Removes an additional attribute from the profile card properties';
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
        force: !!args.options.force
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
    const directoryPropertyName = profileCardPropertyNames.find(n => n.toLowerCase() === args.options.name.toLowerCase());

    const removeProfileCardProperty = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing '${directoryPropertyName}' as a profile card property...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/people/profileCardProperties/${directoryPropertyName}`,
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
        message: `Are you sure you want to remove the profile card property '${directoryPropertyName}'?`
      });
      if (result.continue) {
        await removeProfileCardProperty();
      }
    }
  }
}

export default new TenantPeopleProfileCardPropertyRemoveCommand();