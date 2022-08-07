import { Logger } from '../../../../cli';
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
  pageName: string;
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

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --pageName <pageName>'
      },
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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let pageName: string = args.options.pageName;
    if (args.options.pageName.indexOf('.aspx') < 0) {
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
}

module.exports = new SpoPageControlListCommand();