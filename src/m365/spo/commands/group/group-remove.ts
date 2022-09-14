import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: number;
  name?: string;
  confirm?: boolean;
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
        confirm: (!(!args.options.confirm)).toString()
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
        option: '--confirm'
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
    this.optionSets.push(['id', 'name']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeGroup: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        logger.logToStderr(`Removing group in web at ${args.options.webUrl}...`);
      }

      try {
        let groupId: number | undefined;
        if (args.options.name) {
          const requestOptions: any = {
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
        const requestOptions: any = {
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

    if (args.options.confirm) {
      await removeGroup();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the group ${args.options.id || args.options.name} from web ${args.options.webUrl}?`
      });

      if (result.continue) {
        await removeGroup();
      }
    }
  }
}

module.exports = new SpoGroupRemoveCommand();