import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: number;
  name?: string;
  force?: boolean;
}

class SpoGroupRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_REMOVE;
  }

  public get description(): string {
    return 'Removes group from specific web';
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
        id: (!(!args.options.id)).toString(),
        name: (!(!args.options.name)).toString(),
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--id [id]'
      },
      {
        option: '--name [name]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.id && typeof args.options.id !== 'number') {
          return `${args.options.id} is not a number`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeGroup = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing group in web at ${args.options.webUrl}...`);
      }

      try {
        let groupId: number | undefined;
        if (args.options.name) {
          const requestOptions: CliRequestOptions = {
            url: `${args.options.webUrl}/_api/web/sitegroups/GetByName('${args.options.name}')?$select=Id`,
            headers: {
              accept: 'application/json'
            },
            responseType: 'json'
          };
          const group = await request.get<{ Id: number }>(requestOptions);
          groupId = group.Id;
        }
        else {
          groupId = args.options.id;
        }

        const requestUrl = `${args.options.webUrl}/_api/web/sitegroups/RemoveById(${groupId})`;
        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          method: 'POST',
          headers: {
            'content-length': 0,
            'accept': 'application/json'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
        // REST post call doesn't return anything
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroup();
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to remove the group ${args.options.id || args.options.name} from web ${args.options.webUrl}?`);

      if (result) {
        await removeGroup();
      }
    }
  }
}

export default new SpoGroupRemoveCommand();