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
  fieldId?: string;
  fieldTitle?: string;
  fieldPosition?: string;
  listId?: string;
  listTitle?: string;
  viewId?: string;
  viewTitle?: string;
  webUrl: string;
}

class SpoListViewFieldAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_FIELD_ADD;
  }

  public get description(): string {
    return 'Adds the specified field to list view';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.viewId = typeof args.options.viewId !== 'undefined';
    telemetryProps.viewTitle = typeof args.options.viewTitle !== 'undefined';
    telemetryProps.fieldId = typeof args.options.fieldId !== 'undefined';
    telemetryProps.fieldTitle = typeof args.options.fieldTitle !== 'undefined';
    telemetryProps.fieldPosition = typeof args.options.fieldPosition !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const listSelector: string = args.options.listId ? `(guid'${encodeURIComponent(args.options.listId)}')` : `/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')`;
    let viewSelector: string = '';
    let currentField: { InternalName: string; };

    if (this.verbose) {
      cmd.log(`Getting field ${args.options.fieldId || args.options.fieldTitle}...`);
    }

    this
      .getField(args.options, listSelector)
      .then((field: { InternalName: string; }): Promise<void> => {
        if (this.verbose) {
          cmd.log(`Adding the field ${args.options.fieldId || args.options.fieldTitle} to the view ${args.options.viewId || args.options.viewTitle}...`);
        }

        currentField = field;

        viewSelector = args.options.viewId ? `('${encodeURIComponent(args.options.viewId)}')` : `/GetByTitle('${encodeURIComponent(args.options.viewTitle as string)}')`;
        const postRequestUrl: string = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}/viewfields/addviewfield('${field.InternalName}')`;

        const postRequestOptions: any = {
          url: postRequestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.post(postRequestOptions);
      })
      .then((): Promise<void> => {
        if (typeof args.options.fieldPosition === 'undefined') {
          if (this.debug) {
            cmd.log(`No field position.`);
          }

          return Promise.resolve();
        }

        if (this.debug) {
          cmd.log(`moveField request...`);
          cmd.log(args.options.fieldPosition);
        }

        if (this.verbose) {
          cmd.log(`Moving the field ${args.options.fieldId || args.options.fieldTitle} to the position ${args.options.fieldPosition} from view ${args.options.viewId || args.options.viewTitle}...`);
        }
        const moveRequestUrl: string = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}/viewfields/moveviewfieldto`;

        const moveRequestOptions: any = {
          url: moveRequestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          body: { 'field': currentField.InternalName, 'index': args.options.fieldPosition },
          json: true
        };

        return request.post(moveRequestOptions);
      })
      .then((): void => {
        // REST post call doesn't return anything
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));

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
        description: 'ID of the list where the view is located. Specify `listTitle` or `listId` but not both'
      },
      {
        option: '--listTitle [listTitle]',
        description: 'Title of the list where the view is located. Specify `listTitle` or `listId` but not both'
      },
      {
        option: '--viewId [viewId]',
        description: 'ID of the view to update. Specify `viewTitle` or `viewId` but not both'
      },
      {
        option: '--viewTitle [viewTitle]',
        description: 'Title of the view to update. Specify `viewTitle` or `viewId` but not both'
      },
      {
        option: '--fieldId [fieldId]',
        description: 'ID of the field to add. Specify `fieldId` or `fieldTitle` but not both'
      },
      {
        option: '--fieldTitle [fieldTitle]',
        description: 'The case-sensitive internal name or display name of the field to add. Specify `fieldId` or `fieldTitle` but not both'
      },
      {
        option: '--fieldPosition [fieldPosition]',
        description: 'The zero-based index of the position for the field'
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
          return `${args.options.fieldId} is not a valid GUID`;
        }
      }

      if (args.options.fieldPosition) {
        const position: number = parseInt(args.options.fieldPosition);
        if (isNaN(position)) {
          return `${args.options.fieldPosition} is not a number`;
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
  
    Add field with ID ${chalk.grey('330f29c5-5c4c-465f-9f4b-7903020ae1ce')} to view
    with ID ${chalk.grey('3d760127-982c-405e-9c93-e1f76e1a1110')} of the list with ID
    ${chalk.grey('1f187321-f086-4d3d-8523-517e94cc9df9')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LIST_VIEW_FIELD_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --fieldId 330f29c5-5c4c-465f-9f4b-7903020ae1ce --listId 1f187321-f086-4d3d-8523-517e94cc9df9 --viewId 3d760127-982c-405e-9c93-e1f76e1a1110

    Add field with title ${chalk.grey('Custom field')} to view with title ${chalk.grey('All Documents')}
    of the list with title ${chalk.grey('Documents')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LIST_VIEW_FIELD_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --fieldTitle 'Custom field' --listTitle Documents --viewTitle 'All Documents'
    
    Add field with title ${chalk.grey('Custom field')} at the position ${chalk.grey('0')} to view with title
    ${chalk.grey('All Documents')} of the list with title ${chalk.grey('Documents')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${commands.LIST_VIEW_FIELD_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --fieldTitle 'Custom field' --listTitle Documents --viewTitle 'All Documents' --fieldPosition 0
      `);
  }
}

module.exports = new SpoListViewFieldAddCommand();