import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id: string;
  scope?: string;
  force: boolean;
}

class SpoFeatureEnableCommand extends SpoCommand {
  public get name(): string {
    return commands.FEATURE_ENABLE;
  }

  public get description(): string {
    return 'Enables feature for the specified site or web';
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
        if (args.options.scope) {
          if (['site', 'web'].indexOf(args.options.scope.toLowerCase()) < 0) {
            return `${args.options.scope} is not a valid Feature scope. Allowed values are Site|Web`;
          }
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
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
      logger.logToStderr(`Enabling feature '${args.options.id}' on scope '${scope}' for url '${args.options.webUrl}' (force='${force}')...`);
    }

    const url: string = `${args.options.webUrl}/_api/${scope}/features/add(featureId=guid'${args.options.id}',force=${force})`;
    const requestOptions: any = {
      url: url,
      headers: {
        accept: 'application/json;odata=nometadata'
      }
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFeatureEnableCommand();