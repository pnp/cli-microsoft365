import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Control } from './canvasContent';
import { ClientSidePageProperties } from './ClientSidePageProperties';
import { getControlTypeDisplayName } from './pageMethods';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
}

class SpoPageControlListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_CONTROL_LIST;
  }

  public get description(): string {
    return 'Lists controls on the specific modern page';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'type', 'title'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let pageName: string = args.options.name;
    if (args.options.name.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/SitePages/Pages/GetByUrl('sitepages/${encodeURIComponent(pageName)}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<ClientSidePageProperties>(requestOptions)
      .then((clientSidePage: ClientSidePageProperties): void => {
        const canvasData: Control[] = clientSidePage.CanvasContent1 ? JSON.parse(clientSidePage.CanvasContent1) : [];
        const controls: any[] = canvasData.filter(c => c.position).map(c => {
          return {
            id: c.id,
            type: getControlTypeDisplayName(
              c.controlType || 0
            ),
            title: c.webPartData?.title,
            controlType: c.controlType,
            order: c.position.sectionIndex,
            controlData: {
              ...c
            }
          };
        });

        logger.log(JSON.parse(JSON.stringify(controls)));

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
  
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageControlListCommand();