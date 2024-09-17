import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  spObjectId: string;
}

class EntraOAuth2GrantListCommand extends GraphCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_LIST;
  }

  public get description(): string {
    return 'Lists OAuth2 permission grants for the specified service principal';
  }

  public alias(): string[] | undefined {
    return [aadCommands.OAUTH2GRANT_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['objectId', 'resourceId', 'scope'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --spObjectId <spObjectId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.spObjectId)) {
          return `${args.options.spObjectId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.OAUTH2GRANT_LIST, commands.OAUTH2GRANT_LIST);

    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of OAuth grants for the service principal...`);
    }

    try {
      const res = await odata.getAllItems<any>(`${this.resource}/v1.0/oauth2PermissionGrants?$filter=clientId eq '${formatting.encodeQueryParameter(args.options.spObjectId)}'`);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraOAuth2GrantListCommand();