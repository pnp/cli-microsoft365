import { Logger } from "../../../../cli/Logger.js";
import GlobalOptions from "../../../../GlobalOptions.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { validation } from '../../../../utils/validation.js';
import SpoCommand from "../../../base/SpoCommand.js";
import commands from "../../commands.js";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id: string;
  scope?: string;
  force: boolean;
}

class SpoFeatureDisableCommand extends SpoCommand {
  public get name(): string {
    return commands.FEATURE_DISABLE;
  }

  public get description(): string {
    return 'Disables feature for the specified site or web';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        scope: args.options.scope || 'web',
        force: args.options.force || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web']
      },
      {
        option: '--force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.scope) {
          if (['site', 'web'].indexOf(args.options.scope.toLowerCase()) < 0) {
            return `${args.options.scope} is not a valid Feature scope. Allowed values are Site|Web`;
          }
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('scope', 's');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let scope: string | undefined = args.options.scope;
    let force: boolean = args.options.force;

    if (!scope) {
      scope = "web";
    }
    if (!force) {
      force = false;
    }

    if (this.verbose) {
      await logger.logToStderr(`Disabling feature '${args.options.id}' on scope '${scope}' for url '${args.options.webUrl}' (force='${force}')...`);
    }

    const url: string = `${args.options.webUrl}/_api/${scope}/features/remove(featureId=guid'${args.options.id}',force=${force})`;

    const requestOptions: CliRequestOptions = {
      url: url,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFeatureDisableCommand();