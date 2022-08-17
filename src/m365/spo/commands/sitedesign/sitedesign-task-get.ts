import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteDesignTask } from './SiteDesignTask';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  taskId: string;
}

class SpoSiteDesignTaskGetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_TASK_GET;
  }

  public get description(): string {
    return 'Gets information about the specified site design scheduled for execution';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --taskId <taskId>'
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    spo
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<SiteDesignTask> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTask`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          data: {
            taskId: args.options.taskId
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res: SiteDesignTask): void => {
        if (!res["odata.null"]) {
          logger.log(res);
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoSiteDesignTaskGetCommand();