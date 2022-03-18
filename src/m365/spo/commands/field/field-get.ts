import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil, validation } from '../../../../utils';
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
  id?: string;
  fieldTitle?: string;
}

class SpoFieldGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_GET;
  }

  public get description(): string {
    return 'Retrieves information about the specified list- or site column';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.listUrl = typeof args.options.listUrl !== 'undefined';
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.fieldTitle = typeof args.options.fieldTitle !== 'undefined';
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
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);

      listRestUrl = `GetList('${encodeURIComponent(listServerRelativeUrl)}')/`;
    }

    let fieldRestUrl: string = '';
    if (args.options.id) {
      fieldRestUrl = `/getbyid('${encodeURIComponent(args.options.id)}')`;
    }
    else {
      fieldRestUrl = `/getbyinternalnameortitle('${encodeURIComponent(args.options.fieldTitle as string)}')`;
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/${listRestUrl}fields${fieldRestUrl}`,
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
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--fieldTitle [fieldTitle]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (!args.options.id && !args.options.fieldTitle) {
      return 'Specify id or fieldTitle, one is required';
    }

    if (args.options.id && !validation.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
      return `${args.options.listId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoFieldGetCommand();