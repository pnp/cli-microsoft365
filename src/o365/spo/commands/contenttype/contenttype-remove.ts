import auth from '../../SpoAuth';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate, CommandTypes, CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';

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
    return 'Removes an unused content type';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['id', 'i']
    };
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    //track if using id or name parameter
    telemetryProps.lookupBy = args.options.id ? 'ContentType ID' : 'ContentType Name';
    return telemetryProps;
  }


  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let contentTypeId: string = '';

    const contentTypeIdentifierLabel = args.options.id ? 
    `with id ${args.options.id}` : 
    `with name ${args.options.name}`;
    
    const removeContentType = (): void => {
      auth
        .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
        .then((accessToken: string): Promise<any> => {
          siteAccessToken = accessToken;

          if (this.debug) {
            cmd.log(`Retrieved access token ${accessToken}. Retrieving information about the content type ${contentTypeIdentifierLabel}...`);
          }

          if (args.options.name) {
            if (this.verbose) {
              cmd.log(`Looking up content type id ${contentTypeIdentifierLabel}...`);
            }

            const requestOptions: any = {
              url: `${args.options.webUrl}/_api/web/availableContentTypes?$filter=(Name eq '${encodeURIComponent(args.options.name)}')`,
              headers: {
                authorization: `Bearer ${siteAccessToken}`,
                accept: 'application/json;odata=nometadata'
              },
              json: true
            };

            return request.get(requestOptions);
            
          }
          else {
            return Promise.resolve({"value":[{"StringId":args.options.id}]});
        
          }
        })
        .then((contentTypeIdResult: any): Promise<any> => {
          

          if (this.debug) {
            cmd.log(vorpal.chalk.yellow("Raw Respose:"));
            cmd.log(JSON.stringify(contentTypeIdResult));
          }

          if (contentTypeIdResult && contentTypeIdResult.value && contentTypeIdResult.value.length>0)
          {
            if (this.debug) 
            {
              cmd.log(vorpal.chalk.yellow("Parsed Response:"));
              cmd.log(contentTypeIdResult.value[0].StringId);
            }

            contentTypeId = contentTypeIdResult.value[0].StringId;
          
            //execute delete operation
            const requestOptions: any = {
              url: `${args.options.webUrl}/_api/web/contenttypes('${encodeURIComponent(contentTypeId)}')`,
              method: 'POST',
              headers: {
                authorization: `Bearer ${siteAccessToken}`,
                'X-HTTP-Method': 'DELETE',
                'If-Match': '*',
                'accept': 'application/json;odata=nometadata'
              },
              json: true
            };

            return request.post(requestOptions);
          }
          else {
            return Promise.resolve({"odata.null":true});
          }
         
      })
      .then((res): void => {
          //no response object expected
          if (this.debug) {
            cmd.log("deletion response:");
            cmd.log(JSON.stringify(res));
          }
          
            
          if (res && res["odata.null"] === true) {
            if (this.verbose) {
              cmd.log(`Content type not found`);
            }
            cb(new CommandError(`Content type not found`));
            return;
          }
          else {
            
            cmd.log("DONE");
            
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
        cmd.prompt({ type: 'confirm', name: 'continue', default: false, message: `Are you sure you want to remove the content type ${args.options.id||args.options.name}?`}, (result: { continue: boolean }): void => {
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
        description: 'Absolute URL of the site where the content type exists'
      },
      {
        option: '-i, --id [id]',
        description: 'The ID of the content type to remove'
      },
      {
        option: '-n, --name [name]',
        description: 'The name of the content type to remove if ID is unknown'
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
        return 'Either id or name parameter is required missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To remove a content type, you have to first log in to a SharePoint site
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

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
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --id 0x01007926A45D687BA842B947286090B8F67D
    
    Remove a site content type by Name
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'My Content Type'

    Remove a site content type without prompting for confirmation
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'My Content Type' --confirm
    `);
  }
}

module.exports = new SpoContentTypeRemoveCommand();