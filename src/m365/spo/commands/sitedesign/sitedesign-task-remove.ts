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
  taskId: string;
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
        option: '-i, --taskId <taskId>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.taskId)) {
          return `${args.options.taskId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeSiteDesignTask: () => Promise<void> = async (): Promise<void> => {
      try {
        const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RemoveSiteDesignTask`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          data: {
            taskId: args.options.taskId
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      } 
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };
    if (args.options.confirm) {
      await removeSiteDesignTask();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the site design task ${args.options.taskId}?`
      });
      
      if (result.continue) {
        await removeSiteDesignTask();
      }
    }
  }
}

module.exports = new SpoSiteDesignTaskRemoveCommand();