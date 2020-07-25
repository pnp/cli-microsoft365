import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
    return `${commands.FIELD_GET}`;
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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
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
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
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
        description: 'Absolute URL of the site where the field is located'
      },
      {
        option: '-l, --listTitle [listTitle]',
        description: 'Title of the list where the field is located. Specify only one of listTitle, listId or listUrl'
      },
      {
        option: '--listId [listId]',
        description: 'ID of the list where the field is located. Specify only one of listTitle, listId or listUrl'
      },
      {
        option: '--listUrl [listUrl]',
        description: 'Server- or web-relative URL of the list where the field is located. Specify only one of listTitle, listId or listUrl'
      },
      {
        option: '-i, --id [id]',
        description: 'The ID of the field to retrieve. Specify id or fieldTitle but not both'
      },
      {
        option: '--fieldTitle [fieldTitle]',
        description: 'The display name (case-sensitive) of the field to retrieve. Specify id or fieldTitle but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.id && !args.options.fieldTitle) {
        return 'Specify id or fieldTitle, one is required';
      }

      if (args.options.id && !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.listId && !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new SpoFieldGetCommand();