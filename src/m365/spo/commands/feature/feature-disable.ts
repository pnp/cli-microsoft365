import { Logger } from "../../../../cli";
import GlobalOptions from "../../../../GlobalOptions";
import request from "../../../../request";
import SpoCommand from "../../../base/SpoCommand";
import commands from "../../commands";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  featureId: string;
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
        option: '-u, --url <url>'
      },
      {
        option: '-f, --featureId <id>'
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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let scope: string | undefined = args.options.scope;
    let force: boolean = args.options.force;

    if (!scope) {
      scope = "web";
    }
    if (!force) {
      force = false;
    }

    if (this.verbose) {
      logger.logToStderr(`Disabling feature '${args.options.featureId}' on scope '${scope}' for url '${args.options.url}' (force='${force}')...`);
    }

    const url: string = `${args.options.url}/_api/${scope}/features/remove(featureId=guid'${args.options.featureId}',force=${force})`;

    const requestOptions: any = {
      url: url,
      headers: {
        accept: 'application/json;odata=nometadata'
      }
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));

  }
}

module.exports = new SpoFeatureDisableCommand();