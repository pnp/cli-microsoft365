import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { GetClientSideWebPartsRsp } from './GetClientSideWebPartsRsp';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoWebClientSideWebPartListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_CLIENTSIDEWEBPART_LIST;
  }

  public get description(): string {
    return 'Lists available client-side web parts';
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
      url: `${args.options.webUrl}/_api/web/GetClientSideWebParts`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res: GetClientSideWebPartsRsp = await request.get<GetClientSideWebPartsRsp>(requestOptions);

      const clientSideWebParts: any[] = [];
      res.value.forEach(component => {
        if (component.ComponentType === 1) {
          clientSideWebParts.push({
            Id: component.Id.replace("{", "").replace("}", ""),
            Name: component.Name,
            Title: JSON.parse(component.Manifest).preconfiguredEntries[0].title.default
          });
        }
      });

      if (clientSideWebParts.length === 0 && this.verbose) {
        logger.logToStderr("No client-side web parts available for this site");
      }

      if (clientSideWebParts.length > 0) {
        logger.log(clientSideWebParts);
      }
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoWebClientSideWebPartListCommand();