import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const listSelector: string = args.options.listId ? `(guid'${encodeURIComponent(args.options.listId)}')` : `/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')`;

    const removeFieldFromView: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Getting field ${args.options.fieldId || args.options.fieldTitle}...`);
      }

      this
        .getField(args.options, listSelector)
        .then((field: { InternalName: string; }): Promise<void> => {
          if (this.verbose) {
            cmd.log(`Removing field ${args.options.fieldId || args.options.fieldTitle} from view ${args.options.viewId || args.options.viewTitle}...`);
          }

          const viewSelector: string = args.options.viewId ? `('${encodeURIComponent(args.options.viewId)}')` : `/GetByTitle('${encodeURIComponent(args.options.viewTitle as string)}')`;
          const postRequestUrl: string = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}/viewfields/removeviewfield('${field.InternalName}')`;

          const postRequestOptions: any = {
            url: postRequestUrl,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            json: true
          };

          return request.post(postRequestOptions);
        })
        .then((): void => {
          // REST post call doesn't return anything
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeFieldFromView();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the field ${args.options.fieldId || args.options.fieldTitle} from the view ${args.options.viewId || args.options.viewTitle} from list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}?`,
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
      json: true
    };

    return request.get(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list is located'
      },
      {
        option: '--listId [listId]',
        description: 'ID of the list where the view is located. Specify listTitle or listId but not both'
      },
      {
        option: '--listTitle [listTitle]',
        description: 'Title of the list where the view is located. Specify listTitle or listId but not both'
      },
      {
        option: '--viewId [viewId]',
        description: 'ID of the view to update. Specify viewTitle or viewId but not both'
      },
      {
        option: '--viewTitle [viewTitle]',
        description: 'Title of the view to update. Specify viewTitle or viewId but not both'
      },
      {
        option: '--fieldId [fieldId]',
        description: 'ID of the field to remove. Specify fieldId or fieldTitle but not both'
      },
      {
        option: '--fieldTitle [fieldTitle]',
        description: 'The case-sensitive internal name or display name of the field to remove. Specify fieldId or fieldTitle but not both'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the field from the view'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.listId) {
        if (!Utils.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }
      }

      if (args.options.viewId) {
        if (!Utils.isValidGuid(args.options.viewId)) {
          return `${args.options.viewId} is not a valid GUID`;
        }
      }

      if (args.options.fieldId) {
        if (!Utils.isValidGuid(args.options.fieldId)) {
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
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Remove field with ID ${chalk.grey('330f29c5-5c4c-465f-9f4b-7903020ae1ce')} from view
    with ID ${chalk.grey('3d760127-982c-405e-9c93-e1f76e1a1110')} from the list
    with ID ${chalk.grey('1f187321-f086-4d3d-8523-517e94cc9df9')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LIST_VIEW_FIELD_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --fieldId 330f29c5-5c4c-465f-9f4b-7903020ae1ce --listId 1f187321-f086-4d3d-8523-517e94cc9df9 --viewId 3d760127-982c-405e-9c93-e1f76e1a1110
      
    Remove field with title ${chalk.grey('Custom field')} from view with title ${chalk.grey('Custom view')}
    from the list with title ${chalk.grey('Documents')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LIST_VIEW_FIELD_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --fieldTitle 'Custom field' --listTitle Documents --viewTitle 'Custom view'
      `);
  }
}

module.exports = new SpoListViewFieldRemoveCommand();