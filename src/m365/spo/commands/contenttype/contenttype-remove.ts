import { Cli, Logger } from '../../../../cli';
import { CommandError } from '../../../../Command';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let contentTypeId: string = '';

    const contentTypeIdentifierLabel: string = args.options.id ?
      `with id ${args.options.id}` :
      `with name ${args.options.name}`;

    const removeContentType = (): void => {
      ((): Promise<any> => {
        if (this.debug) {
          logger.logToStderr(`Retrieving information about the content type ${contentTypeIdentifierLabel}...`);
        }

        if (args.options.id) {
          return Promise.resolve({ "value": [{ "StringId": args.options.id }] });
        }

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

        return request.get(requestOptions);
      })()
        .then((contentTypeIdResult: { value: { StringId: string }[] }): Promise<any> => {
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

            return request.post(requestOptions);
          }
          else {
            return Promise.resolve({ "odata.null": true });
          }
        })
        .then((res): void => {
          if (res && res["odata.null"] === true) {
            cb(new CommandError(`Content type not found`));
            return;
          }

          cb();
        }, (err: any): void => {
          this.handleRejectedODataJsonPromise(err, logger, cb);
        });
    };

    if (args.options.confirm) {
      removeContentType();
    }
    else {
      Cli.prompt({ type: 'confirm', name: 'continue', default: false, message: `Are you sure you want to remove the content type ${args.options.id || args.options.name}?` }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeContentType();
        }
      });
    }
  }
}

module.exports = new SpoContentTypeRemoveCommand();