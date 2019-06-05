import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate, CommandTypes, CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    return `${commands.CONTENTTYPE_REMOVE}`;
  }

  public get description(): string {
    return 'Deletes site content type';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['id', 'i']
    };
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let contentTypeId: string = '';

    const contentTypeIdentifierLabel: string = args.options.id ?
      `with id ${args.options.id}` :
      `with name ${args.options.name}`;

    const removeContentType = (): void => {
      ((): Promise<any> => {
        if (this.debug) {
          cmd.log(`Retrieving information about the content type ${contentTypeIdentifierLabel}...`);
        }

        if (args.options.id) {
          return Promise.resolve({ "value": [{ "StringId": args.options.id }] });
        }

        if (this.verbose) {
          cmd.log(`Looking up the ID of content type ${contentTypeIdentifierLabel}...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/availableContentTypes?$filter=(Name eq '${encodeURIComponent(args.options.name as string)}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
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
              json: true
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
          else {
            if (this.verbose) {
              cmd.log(vorpal.chalk.green('DONE'));
            }
          }

          cb();
        }, (err: any): void => {
          this.handleRejectedODataJsonPromise(err, cmd, cb);
        });
    }

    if (args.options.confirm) {
      removeContentType();
    }
    else {
      cmd.prompt({ type: 'confirm', name: 'continue', default: false, message: `Are you sure you want to remove the content type ${args.options.id || args.options.name}?` }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeContentType();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site where the content type is located'
      },
      {
        option: '-i, --id [id]',
        description: 'The ID of the content type to remove'
      },
      {
        option: '-n, --name [name]',
        description: 'The name of the content type to remove'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removal of the content type'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
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
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    If the specified content type is in use by a list and cannot be removed, 
    you will be returned the error:
    ${chalk.grey('Another site or list is still using this content type.')}
    SharePoint will not allow a content type to be removed unless any 
    dependent objects are also emptied from the recycle bin including 
    the second-stage recycle bin.

    The content type you wish to remove can be selected by the ID or Name 
    of the content type. Either ID or Name parameter must be specified.

  Examples:
  
    Remove a site content type by ID
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --id 0x01007926A45D687BA842B947286090B8F67D
    
    Remove a site content type by Name
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'My Content Type'

    Remove a site content type without prompting for confirmation
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'My Content Type' --confirm
    `);
  }
}

module.exports = new SpoContentTypeRemoveCommand();