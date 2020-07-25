import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandTypes,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  contentTypeId: string;
  confirm?: boolean;
}

class SpoListContentTypeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_CONTENTTYPE_REMOVE;
  }

  public get description(): string {
    return 'Removes content type from list';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['contentTypeId', 'c']
    };
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeContentTypeFromList: () => void = (): void => {
      if (this.verbose) {
        const list: string = args.options.listId ? encodeURIComponent(args.options.listId as string) : encodeURIComponent(args.options.listTitle as string);
        cmd.log(`Removing content type ${args.options.contentTypeId} from list ${list} in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = '';

      if (args.options.listId) {
        requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')/ContentTypes('${encodeURIComponent(args.options.contentTypeId)}')`;
      }
      else {
        requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')/ContentTypes('${encodeURIComponent(args.options.contentTypeId)}')`;
      }

      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };

      request
        .post(requestOptions)
        .then((): void => {
          // REST post call doesn't return anything
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeContentTypeFromList();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the content type ${args.options.contentTypeId} from the list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeContentTypeFromList();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list from which to remove the content type. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list from which to remove the content type. Specify listId or listTitle but not both'
      },
      {
        option: '-c, --contentTypeId <contentTypeId>',
        description: 'ID of the content type to remove from the list'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the content type from the list'
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

      if (args.options.listId) {
        if (!Utils.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }
      }

      if (args.options.listId && args.options.listTitle) {
        return 'Specify listId or listTitle, but not both';
      }

      if (!args.options.listId && !args.options.listTitle) {
        return 'Specify listId or listTitle, one is required';
      }

      return true;
    };
  }
}

module.exports = new SpoListContentTypeRemoveCommand();