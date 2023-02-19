import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoRoleDefinitionListCommand extends SpoCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_LIST;
  }

  public get description(): string {
    return 'Gets list of role definitions for the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Name'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Getting role definitions list from ${args.options.webUrl}...`);
    }

    try {
      const res = await odata.getAllItems<any>(`${args.options.webUrl}/_api/web/roledefinitions`);
      const response = formatting.setFriendlyPermissions(res);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoRoleDefinitionListCommand();