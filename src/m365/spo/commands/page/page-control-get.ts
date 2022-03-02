import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Control } from './canvasContent';
import { ClientSidePageProperties } from './ClientSidePageProperties';
import { getControlTypeDisplayName } from './pageMethods';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  name: string;
  webUrl: string;
}

class SpoPageControlGetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_CONTROL_GET;
  }

  public get description(): string {
    return 'Gets information about the specific control on a modern page';
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
        const control: Control | undefined = canvasData.find(c => c.id?.toLowerCase() === args.options.id.toLowerCase());

        if (control) {
          const controlData = {
            id: control.id,
            type: getControlTypeDisplayName(
              control.controlType || 0
            ),
            title: control.webPartData?.title,
            controlType: control.controlType,
            order: control.position.sectionIndex,
            controlData: {
              ...control
            }
          };

          logger.log(controlData);
        }
        else {
          if (this.verbose) {
            logger.logToStderr(`Control with ID ${args.options.id} not found on page ${args.options.name}`);
          }
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>'
      },
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
    if (!validation.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageControlGetCommand();