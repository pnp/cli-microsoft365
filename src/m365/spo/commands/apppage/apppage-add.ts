import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  title: string;
  webPartData: string;
  addToQuickLaunch: boolean;
}

class SpoAppPageAddCommand extends SpoCommand {
  public get name(): string {
    return commands.APPPAGE_ADD;
  }

  public get description(): string {
    return 'Creates a single-part app page';
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
        addToQuickLaunch: args.options.addToQuickLaunch
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-t, --title <title>'
      },
      {
        option: '-d, --webPartData <webPartData>'
      },
      {
        option: '--addToQuickLaunch'
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
    const createPageRequestOptions: any = {
      url: `${args.options.webUrl}/_api/sitepages/Pages/CreateAppPage`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        webPartDataAsJson: args.options.webPartData
      }
    };

    try {
      const page = await request.post<{ value: string }>(createPageRequestOptions);

      const pageUrl: string = page.value;

      let requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/${pageUrl}')?$expand=ListItemAllFields`,
        headers: {
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const file = await request.get<{ ListItemAllFields: { Id: string; }; }>(requestOptions);

      requestOptions = {
        url: `${args.options.webUrl}/_api/sitepages/Pages/UpdateAppPage`,
        headers: {
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          pageId: file.ListItemAllFields.Id,
          webPartDataAsJson: args.options.webPartData,
          title: args.options.title,
          includeInNavigation: args.options.addToQuickLaunch
        }
      };

      const res = await request.post(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoAppPageAddCommand();