import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  name?: string;
  confirm?: boolean;
}

class SpoContentTypeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_REMOVE;
  }

  public get description(): string {
    return 'Deletes site content type';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
    
        if (!args.options.id && !args.options.name) {
          return 'Specify either the id or the name';
        }
    
        if (args.options.id && args.options.name) {
          return 'Specify either the id or the name but not both';
        }
    
        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'i');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let contentTypeId: string = '';

    const contentTypeIdentifierLabel: string = args.options.id ?
      `with id ${args.options.id}` :
      `with name ${args.options.name}`;

    const removeContentType: () => Promise<void> = async (): Promise<void> => {
      try {
        if (this.debug) {
          logger.logToStderr(`Retrieving information about the content type ${contentTypeIdentifierLabel}...`);
        }

        let contentTypeIdResult: { value: { StringId: string }[] };
        if (args.options.id) {
          contentTypeIdResult = { "value": [{ "StringId": args.options.id }] };
        }
        else {
          if (this.verbose) {
            logger.logToStderr(`Looking up the ID of content type ${contentTypeIdentifierLabel}...`);
          }
  
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/availableContentTypes?$filter=(Name eq '${encodeURIComponent(args.options.name as string)}')`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };
  
          contentTypeIdResult = await request.get<{ value: { StringId: string }[] }>(requestOptions);
        }

        let res: any;
        if (contentTypeIdResult &&
          contentTypeIdResult.value &&
          contentTypeIdResult.value.length > 0) {
          contentTypeId = contentTypeIdResult.value[0].StringId;

          //execute delete operation
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/contenttypes('${encodeURIComponent(contentTypeId)}')`,
            headers: {
              'X-HTTP-Method': 'DELETE',
              'If-Match': '*',
              'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          res = await request.post<any>(requestOptions);
        }
        else {
          res = { "odata.null": true };
        }

        if (res && res["odata.null"] === true) {
          throw `Content type not found`;
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeContentType();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({ type: 'confirm', name: 'continue', default: false, message: `Are you sure you want to remove the content type ${args.options.id || args.options.name}?` });

      if (result.continue) {
        await removeContentType();
      }
    }
  }
}

module.exports = new SpoContentTypeRemoveCommand();