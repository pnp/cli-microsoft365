import { Logger } from '../../../../cli';
import { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface  Options extends GlobalOptions {
  webUrl: string;
  listTitle?: string;
  id?: string;
  name?: string;
}

class SpoContentTypeGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_GET;
  }

  public get description(): string {
    return 'Retrieves information about the specified list or site content type';
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
    this.#initOptionSets();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listTitle: typeof args.options.listTitle !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listTitle [listTitle]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
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
    
        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'i');
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'name']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let requestUrl: string = `${args.options.webUrl}/_api/web/${(args.options.listTitle ? `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/` : '')}contenttypes`;

    if (args.options.id) {
      requestUrl += `('${encodeURIComponent(args.options.id)}')`;
    }
    else if (args.options.name) {
      requestUrl += `?$filter=Name eq '${encodeURIComponent(args.options.name)}'`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        let errorMessage: string = '';

        if (args.options.name) {
          if (res.value.length === 0) {
            errorMessage = `Content type with name ${args.options.name} not found`;
          }
          else{
            res = res.value[0];
          }
        }

        if (args.options.id && res['odata.null'] === true) {
          errorMessage = `Content type with ID ${args.options.id} not found`;
        }

        if (errorMessage) {
          cb(new CommandError(errorMessage));
          return;
        }

        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoContentTypeGetCommand();