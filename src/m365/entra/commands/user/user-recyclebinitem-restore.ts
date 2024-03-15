import { User } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class EntraUserRecycleBinItemRestoreCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores a user from the tenant recycle bin';
  }

  public alias(): string[] | undefined {
    return [aadCommands.USER_RECYCLEBINITEM_RESTORE];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.USER_RECYCLEBINITEM_RESTORE, commands.USER_RECYCLEBINITEM_RESTORE);

    if (this.verbose) {
      await logger.logToStderr(`Restoring user with id ${args.options.id} from the recycle bin.`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/directory/deletedItems/${args.options.id}/restore`,
        headers: {
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };
      const user = await request.post<User>(requestOptions);
      await logger.log(user);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserRecycleBinItemRestoreCommand();