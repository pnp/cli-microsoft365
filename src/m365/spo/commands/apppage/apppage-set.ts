import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { urlUtil } from '../../../../utils/urlUtil';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  name: string;
  webPartData: string;
}

class SpoAppPageSetCommand extends SpoCommand {
  public get name(): string {
    return commands.APPPAGE_SET;
  }

  public get description(): string {
    return 'Updates the single-part app page';
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
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-d, --webPartData <webPartData>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        try {
          JSON.parse(args.options.webPartData);
        }
        catch (e) {
          return `Specified webPartData is not a valid JSON string. Error: ${e}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/sitepages/Pages/UpdateFullPageApp`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        serverRelativeUrl: `${urlUtil.getServerRelativePath(args.options.webUrl, '')}/SitePages/${args.options.name}`,
        webPartDataAsJson: args.options.webPartData
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
module.exports = new SpoAppPageSetCommand();