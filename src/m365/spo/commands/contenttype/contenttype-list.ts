import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  category?: string;
}

class SpoContentTypeListCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_LIST;
  }

  public get description(): string {
    return 'Lists all available content types in the specified site';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let requestUrl: string = `${args.options.webUrl}/_api/web/ContentTypes`;

    if (args.options.category){
      requestUrl += `?$filter=Group eq '${args.options.category}'`;
    }

    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        logger.log(res.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-c, --category [category]'
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

    return true;
  }
}

module.exports = new SpoContentTypeListCommand();