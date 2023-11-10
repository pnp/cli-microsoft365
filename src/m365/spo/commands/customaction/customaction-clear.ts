import chalk from 'chalk';
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
  scope?: string;
  force?: boolean;
}

class SpoCustomActionClearCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_CLEAR;
  }

  public get description(): string {
    return 'Deletes all custom actions in the collection';
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
        scope: args.options.scope || 'All',
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
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (typeof isValidUrl === 'string') {
          return isValidUrl;
        }

        if (args.options.scope &&
          args.options.scope !== 'Site' &&
          args.options.scope !== 'Web' &&
          args.options.scope !== 'All') {
          return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearCustomActions = async (): Promise<void> => {
      try {
        if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
          await this.clearScopedCustomActions(args.options);
        }
        else {
          await this.clearAllScopes(args.options);
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.force) {
      await clearCustomActions();
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to clear all the user custom actions with scope ${chalk.yellow(args.options.scope || 'All')}?` });

      if (result) {
        await clearCustomActions();
      }
    }
  }

  private clearScopedCustomActions(options: Options): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/${options.scope}/UserCustomActions/clear`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  /**
   * Clear request with `web` scope is send first. 
   * Another clear request is send with `site` scope after.
   */
  private async clearAllScopes(options: Options): Promise<void> {
    options.scope = "Web";

    await this.clearScopedCustomActions(options);

    options.scope = "Site";

    await this.clearScopedCustomActions(options);
  }
}

export default new SpoCustomActionClearCommand();