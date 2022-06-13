import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  fieldId?: string;
  fieldTitle?: string;
  fieldPosition: string;
  listId?: string;
  listTitle?: string;
  viewId?: string;
  viewTitle?: string;
  webUrl: string;
}

class SpoListViewFieldSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_FIELD_SET;
  }

  public get description(): string {
    return 'Updates existing column in an existing view (eg. move to a specific position).';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.viewId = typeof args.options.viewId !== 'undefined';
    telemetryProps.viewTitle = typeof args.options.viewTitle !== 'undefined';
    telemetryProps.fieldId = typeof args.options.fieldId !== 'undefined';
    telemetryProps.fieldTitle = typeof args.options.fieldTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const listSelector: string = args.options.listId ? `(guid'${formatting.encodeQueryParameter(args.options.listId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
    const viewSelector: string = args.options.viewId ? `('${formatting.encodeQueryParameter(args.options.viewId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle as string)}')`;

    if (this.verbose) {
      logger.logToStderr(`Getting field ${args.options.fieldId || args.options.fieldTitle}...`);
    }

    this
      .getField(args.options, listSelector)
      .then((field: { InternalName: string; }): Promise<void> => {
        if (this.verbose) {
          logger.logToStderr(`Moving the field ${args.options.fieldId || args.options.fieldTitle} in view ${args.options.viewId || args.options.viewTitle} to position ${args.options.fieldPosition}...`);
        }

        const moveRequestUrl: string = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}/viewfields/moveviewfieldto`;

        const moveRequestOptions: any = {
          url: moveRequestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          data: {
            field: field.InternalName,
            index: args.options.fieldPosition
          },
          responseType: 'json'
        };

        return request.post(moveRequestOptions);
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));

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
        option: '--fieldPosition <fieldPosition>'
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
        return `${args.options.fieldId} is not a valid GUID`;
      }
    }

    const position: number = parseInt(args.options.fieldPosition);
    if (isNaN(position)) {
      return `${args.options.fieldPosition} is not a number`;
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

module.exports = new SpoListViewFieldSetCommand();