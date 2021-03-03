import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoPageListCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_LIST;
  }

  public get description(): string {
    return 'Lists all modern pages in the given site';
  }

  public defaultProperties(): string[] | undefined {
    return ['FileName', 'Title'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving client-side pages...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/sitepages/pages?$orderby=Title`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ value: any[] }>(requestOptions)
      .then((res: { value: any[] }): void => {
        if (res.value && res.value.length > 0) {
          logger.log(res.value);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageListCommand();