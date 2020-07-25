import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate, CommandTypes, CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
          cmd.log(chalk.green('DONE'));
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
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoContentTypeGetCommand();