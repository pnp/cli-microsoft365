import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class AadM365GroupRenewCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_RENEW;
  }

  public get description(): string {
    return `Renews Microsoft 365 group's expiration`;
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
    if (this.verbose) {
      await logger.logToStderr(`Renewing Microsoft 365 group's expiration: ${args.options.id}...`);
    }

    try {
      const isUnifiedGroup = await aadGroup.isUnifiedGroup(args.options.id);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${args.options.id}' is not a Microsoft 365 group.`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${args.options.id}/renew/`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadM365GroupRenewCommand();