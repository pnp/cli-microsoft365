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
  listTitle?: string;
  id: string;
}

class SpoContentTypeGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CONTENTTYPE_GET}`;
  }

  public get description(): string {
    return 'Retrieves information about the specified list or site content type';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['id', 'i']
    };
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/${(args.options.listTitle ? `lists/getByTitle('${encodeURIComponent(args.options.listTitle)}')/` : '')}contenttypes('${encodeURIComponent(args.options.id)}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (res['odata.null'] === true) {
          cb(new CommandError(`Content type with ID ${args.options.id} not found`));
          return;
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
        description: 'Absolute URL of the site where the content type is located'
      },
      {
        option: '-l, --listTitle [listTitle]',
        description: 'Title of the list where the content type is located (if it is a list content type)'
      },
      {
        option: '-i, --id <id>',
        description: 'The ID of the content type to retrieve'
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

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    If no content type with the specified is found in the site or the list, you
    will get the ${chalk.grey('Content type with ID 0x010012 not found')} error.

  Examples:
  
    Retrieve site content type
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --id 0x0100558D85B7216F6A489A499DB361E1AE2F
    
    Retrieve list content type
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --id 0x0100558D85B7216F6A489A499DB361E1AE2F
`);
  }
}

module.exports = new SpoContentTypeGetCommand();