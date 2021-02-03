import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.addToQuickLaunch = args.options.addToQuickLaunch;
    return telemetryProps;
  }

  public get description(): string {
    return 'Creates a single-part app page';
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
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${Utils.getServerRelativeSiteUrl(args.options.webUrl)}/${pageUrl}')?$expand=ListItemAllFields`,
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    try {
      JSON.parse(args.options.webPartData);
    }
    catch (e) {
      return `Specified webPartData is not a valid JSON string. Error: ${e}`;
    }

    return true;
  }
}

module.exports = new SpoAppPageAddCommand();