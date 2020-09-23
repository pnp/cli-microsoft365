import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import { CommandError, CommandOption, CommandTypes } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
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

        logger.log(res);

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoContentTypeGetCommand();