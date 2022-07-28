import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  viewId?: string;
  viewTitle?: string;
  fieldId?: string;
  fieldTitle?: string;
  confirm?: boolean;
}

class SpoListViewFieldRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_FIELD_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified field from list view';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.viewId = typeof args.options.viewId !== 'undefined';
    telemetryProps.viewTitle = typeof args.options.viewTitle !== 'undefined';
    telemetryProps.fieldId = typeof args.options.fieldId !== 'undefined';
    telemetryProps.fieldTitle = typeof args.options.fieldTitle !== 'undefined';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const listSelector: string = args.options.listId ? `(guid'${formatting.encodeQueryParameter(args.options.listId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;

    const removeFieldFromView: () => void = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Getting field ${args.options.fieldId || args.options.fieldTitle}...`);
      }

      this
        .getField(args.options, listSelector)
        .then((field: { InternalName: string; }): Promise<void> => {
          if (this.verbose) {
            logger.logToStderr(`Removing field ${args.options.fieldId || args.options.fieldTitle} from view ${args.options.viewId || args.options.viewTitle}...`);
          }

          const viewSelector: string = args.options.viewId ? `('${formatting.encodeQueryParameter(args.options.viewId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle as string)}')`;
          const postRequestUrl: string = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}/viewfields/removeviewfield('${field.InternalName}')`;

          const postRequestOptions: any = {
            url: postRequestUrl,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          return request.post(postRequestOptions);
        })
        .then((): void => {
          // REST post call doesn't return anything
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeFieldFromView();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the field ${args.options.fieldId || args.options.fieldTitle} from the view ${args.options.viewId || args.options.viewTitle} from list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeFieldFromView();
        }
      });
    }
  }

  private getField(options: Options, listSelector: string): Promise<{ InternalName: string; }> {
    const fieldSelector: string = options.fieldId ? `/getbyid('${encodeURIComponent(options.fieldId)}')` : `/getbyinternalnameortitle('${encodeURIComponent(options.fieldTitle as string)}')`;
    const getRequestUrl: string = `${options.webUrl}/_api/web/lists${listSelector}/fields${fieldSelector}`;

    const requestOptions: any = {
      url: getRequestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--viewId [viewId]'
      },
      {
        option: '--viewTitle [viewTitle]'
      },
      {
        option: '--fieldId [fieldId]'
      },
      {
        option: '--fieldTitle [fieldTitle]'
      },
      {
        option: '--confirm'
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

    if (args.options.listId) {
      if (!validation.isValidGuid(args.options.listId)) {
        return `${args.options.listId} is not a valid GUID`;
      }
    }

    if (args.options.viewId) {
      if (!validation.isValidGuid(args.options.viewId)) {
        return `${args.options.viewId} is not a valid GUID`;
      }
    }

    if (args.options.fieldId) {
      if (!validation.isValidGuid(args.options.fieldId)) {
        return `${args.options.viewId} is not a valid GUID`;
      }
    }

    if (args.options.listId && args.options.listTitle) {
      return 'Specify listId or listTitle, but not both';
    }

    if (!args.options.listId && !args.options.listTitle) {
      return 'Specify listId or listTitle, one is required';
    }

    if (args.options.viewId && args.options.viewTitle) {
      return 'Specify viewId or viewTitle, but not both';
    }

    if (!args.options.viewId && !args.options.viewTitle) {
      return 'Specify viewId or viewTitle, one is required';
    }

    if (args.options.fieldId && args.options.fieldTitle) {
      return 'Specify fieldId or fieldTitle, but not both';
    }

    if (!args.options.fieldId && !args.options.fieldTitle) {
      return 'Specify fieldId or fieldTitle, one is required';
    }

    return true;
  }
}

module.exports = new SpoListViewFieldRemoveCommand();