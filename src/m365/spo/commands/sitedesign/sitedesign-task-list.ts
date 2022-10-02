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
}

class SpoSiteDesignTaskListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_TASK_LIST;
  }

  public get description(): string {
    return 'Lists site designs scheduled for execution on the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['ID', 'SiteDesignID', 'LogonName'];
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
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTasks`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res: { value: any[] } = await request.post<{ value: any[] }>(requestOptions);
      logger.log(res.value);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteDesignTaskListCommand();