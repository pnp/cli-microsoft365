import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteDesignTask } from './SiteDesignTask';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
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
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTask`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: {
          taskId: args.options.id
        },
        responseType: 'json'
      };

      const res: SiteDesignTask = await request.post(requestOptions);
      if (!res["odata.null"]) {
        logger.log(res);
      }
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteDesignTaskGetCommand();