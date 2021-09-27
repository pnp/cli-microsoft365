import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;  
}

class SpoFieldListCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_LIST;
  }

  public get description(): string {
    return 'Retrieves fields for a given list or site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.listUrl = typeof args.options.listUrl !== 'undefined';    
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let listRestUrl: string = '';

    if (args.options.listId) {
      listRestUrl = `lists(guid'${encodeURIComponent(args.options.listId)}')/`;
    }
    else if (args.options.listTitle) {
      listRestUrl = `lists/getByTitle('${encodeURIComponent(args.options.listTitle as string)}')/`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = Utils.getServerRelativePath(args.options.webUrl, args.options.listUrl);

      listRestUrl = `GetList('${encodeURIComponent(listServerRelativeUrl)}')/`;
    }
    

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/${listRestUrl}fields`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
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
        option: '--listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
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

    if (args.options.listId && !Utils.isValidGuid(args.options.listId)) {
      return `${args.options.listId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoFieldListCommand();