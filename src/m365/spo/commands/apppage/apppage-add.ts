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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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

    request
      .post<{ value: string }>(createPageRequestOptions)
      .then((page: { value: string }): Promise<{ ListItemAllFields: { Id: string; }; }> => {
        const pageUrl: string = page.value;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/${pageUrl}')?$expand=ListItemAllFields`,
          headers: {
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get<{ ListItemAllFields: { Id: string; }; }>(requestOptions);
      })
      .then((file: { ListItemAllFields: { Id: string; }; }): Promise<any> => {
        const requestOptions: any = {
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

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoAppPageAddCommand();