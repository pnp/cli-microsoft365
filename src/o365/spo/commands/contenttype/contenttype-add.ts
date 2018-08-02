import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandOption, CommandValidate, CommandTypes
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listTitle?: string;
  name: string;
  id: string;
  description?: string;
  group?: string;
}

class SpoContentTypeAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CONTENTTYPE_ADD}`;
  }

  public get description(): string {
    return 'Adds a new list or site content type';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['id', 'i']
    };
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        const payload: any = {
          Name: args.options.name,
          Id: { StringValue: '' + args.options.id }
        };
        if (args.options.description) {
          payload.Description = args.options.description;
        }
        if (args.options.group) {
          payload.Group = args.options.group;
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/${(args.options.listTitle ? `lists/getByTitle('${encodeURIComponent(args.options.listTitle)}')/` : '')}contenttypes`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata'
          }),
          body: payload,
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(res);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site where the content type should be created'
      },
      {
        option: '-l, --listTitle [listTitle]',
        description: 'Title of the list where the content type should be created (if it should be created as a list content type)'
      },
      {
        option: '-i, --id <id>',
        description: 'The ID of the content type. Determines the parent content type'
      },
      {
        option: '-n, --name <name>',
        description: 'The name of the content type'
      },
      {
        option: '-d, --description [description]',
        description: 'The description of the content type'
      },
      {
        option: '-g, --group [group]',
        description: 'The group with which the content type should be associated'
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

      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site
    using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To create a content type, you have to first connect to a SharePoint site
    using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    If the specified content type already exists, you will get a
    ${chalk.grey('A duplicate content type "Your Content Type" was found.')} error.

    The ID of the content type specifies the parent content type from which this
    content type inherits.

  Examples:
  
    Create a site content type that inherits from the List item content type
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'PnP Alert' --id 0x01007926A45D687BA842B947286090B8F67D --group 'PnP Content Types'
    
    Create a list content type that inherits from the List item content type
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Alerts --name 'PnP Alert' --id 0x01007926A45D687BA842B947286090B8F67D

  More information:

    Content Type IDs
      https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/aa543822(v%3Doffice.14)
`);
  }
}

module.exports = new SpoContentTypeAddCommand();