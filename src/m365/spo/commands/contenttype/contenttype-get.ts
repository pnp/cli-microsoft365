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
  id?: string;
  contenttypeTitle?: string;
}

class SpoContentTypeGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_GET;
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

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/${(args.options.listTitle ? `lists/getByTitle('${encodeURIComponent(args.options.listTitle)}')/` : '')}contenttypes('${encodeURIComponent(args.options.id)}')`;
    }
    else if (args.options.contenttypeTitle) {
      requestUrl = `${args.options.webUrl}/_api/web/${(args.options.listTitle ? `lists/getByTitle('${encodeURIComponent(args.options.listTitle)}')/` : '')}contenttypes?$filter=Name eq '${encodeURIComponent(args.options.contenttypeTitle)}'`;
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
        if (res['odata.null'] === true) {
          if(args.options.id){
            cb(new CommandError(`Content type with ID ${args.options.id} not found`));
          }
          if(args.options.contenttypeTitle){
            cb(new CommandError(`Content type with title ${args.options.contenttypeTitle} not found`));
          }
          return;
        }

        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
        option: '-c, --contenttypeTitle [contenttypeTitle]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.id && args.options.contenttypeTitle) {
      return 'Specify id or content type title, but not both';
    }

    if (!args.options.id && !args.options.contenttypeTitle) {
      return 'Specify id or content type title, one is required';
    }

    return true;
  }
}

module.exports = new SpoContentTypeGetCommand();