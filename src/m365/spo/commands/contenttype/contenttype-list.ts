import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
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
  category?: string;
}

class SpoContentTypeListCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_LIST;
  }

  public get description(): string {
    return 'Lists all available content types in the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['StringId', 'Name', 'Hidden', 'ReadOnly', 'Sealed'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let requestUrl: string = `${args.options.webUrl}/_api/web/ContentTypes`;

    if (args.options.category){
      requestUrl += `?$filter=Group eq '${encodeURIComponent(args.options.category as string)}'`;
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
    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoContentTypeListCommand();