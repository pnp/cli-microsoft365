import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class EntraGroupSettingGetCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUPSETTING_GET;
  }

  public get description(): string {
    return 'Gets information about the particular group setting';
  }

  public alias(): string[] | undefined {
    return [aadCommands.GROUPSETTING_GET];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraGroupSettingGetCommand();