import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Control } from './canvasContent.js';
import { ClientSidePageProperties } from './ClientSidePageProperties.js';
import { getControlTypeDisplayName } from './pageMethods.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  pageName: string;
  webUrl: string;
}

class SpoPageControlGetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_CONTROL_GET;
  }

  public get description(): string {
    return 'Gets information about the specific control on a modern page';
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
      },
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
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let pageName: string = args.options.pageName;
    if (args.options.pageName.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/SitePages/Pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageName)}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const clientSidePage = await request.get<ClientSidePageProperties>(requestOptions);

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

        await logger.log(controlData);
      }
      else {
        if (this.verbose) {
          await logger.logToStderr(`Control with ID ${args.options.id} not found on page ${args.options.pageName}`);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageControlGetCommand();