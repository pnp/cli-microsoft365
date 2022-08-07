import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  confirm?: boolean;
}

class SpoSiteDesignTaskRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_TASK_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified site design scheduled for execution';
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
        confirm: args.options.confirm || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--confirm'
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeSiteDesignTask: () => void = (): void => {
      spo
        .getSpoUrl(logger, this.debug)
        .then((spoUrl: string): Promise<any> => {
          const requestOptions: any = {
            url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RemoveSiteDesignTask`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            data: {
              taskId: args.options.id
            },
            responseType: 'json'
          };

          return request.post(requestOptions);
        })
        .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };
    if (args.options.confirm) {
      removeSiteDesignTask();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the site design task ${args.options.id}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeSiteDesignTask();
        }
      });
    }
  }
}

module.exports = new SpoSiteDesignTaskRemoveCommand();