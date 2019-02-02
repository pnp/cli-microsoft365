import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';

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
  fieldPosition: number;
}

class SpoListViewFieldAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_FIELD_ADD;
  }

  public get description(): string {
    return 'Add the specified field to list view';
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
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    
    const listSelector: string = args.options.listId ? `(guid'${encodeURIComponent(args.options.listId)}')` : `/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')`;
    let viewSelector: string = '';
    let currentField: any;

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Getting field ${args.options.fieldId || args.options.fieldTitle}...`);
        }

        return this.getField(args.options, listSelector, siteAccessToken, cmd, this.debug);
      })
      .then((field: any): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`getField response...`);
          cmd.log(field);
        }

        if (this.verbose) {
          cmd.log(`Adding the field ${args.options.fieldId || args.options.fieldTitle} to the view ${args.options.viewId || args.options.viewTitle}...`);
        }

        /* Current field backup */
        currentField = field;

        viewSelector = args.options.viewId ? `('${encodeURIComponent(args.options.viewId)}')` : `/GetByTitle('${encodeURIComponent(args.options.viewTitle as string)}')`;
        const postRequestUrl: string = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}/viewfields/addviewfield('${field.InternalName}')`;

        const postRequestOptions: any = {
          url: postRequestUrl,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(postRequestOptions);
          cmd.log('');
        }

        return request.post(postRequestOptions);
      })
      .then((): request.RequestPromise | void => {
        if (args.options.fieldPosition !== undefined) {
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
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'accept': 'application/json;odata=nometadata'
            }),
            body: { 'field': currentField.InternalName, 'index': args.options.fieldPosition },
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(moveRequestOptions);
            cmd.log('');
          }

          return request.post(moveRequestOptions);
        }
        if (this.debug) {
          cmd.log(`No field position.`);
        }
      })
      .then((r: any): void => {
        // REST post call doesn't return anything
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));

  }

  protected getField(options: any, listSelector: string, siteAccessToken: string, cmd: CommandInstance, debug: boolean): request.RequestPromise {
    const fieldSelector = options.fieldId ? `/getbyid('${encodeURIComponent(options.fieldId)}')` : `/getbyinternalnameortitle('${encodeURIComponent(options.fieldTitle as string)}')`;
    const getRequestUrl: string = `${options.webUrl}/_api/web/lists${listSelector}/fields${fieldSelector}`;

    const requestOptions: any = {
      url: getRequestUrl,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${siteAccessToken}`,
        'accept': 'application/json;odata=nometadata'
      }),
      json: true
    };

    if (debug) {
      cmd.log('Executing web request...');
      cmd.log(requestOptions);
      cmd.log('');
    }

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
        description: 'ID of the field to add. Specify fieldId or fieldTitle but not both'
      },
      {
        option: '--fieldTitle [fieldTitle]',
        description: 'The case-sensitive internal name or display name of the field to add. Specify fieldId or fieldTitle but not both'
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
      `  ${chalk.yellow('Important:')} before using this command, log in to SharePoint,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To remove field from a list view, you have to first log in to SharePoint
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Add field with ID ${chalk.grey('330f29c5-5c4c-465f-9f4b-7903020ae1ce')} to view with ID ${chalk.grey('3d760127-982c-405e-9c93-e1f76e1a1110')} from the list with ID ${chalk.grey('1f187321-f086-4d3d-8523-517e94cc9df9')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_VIEW_FIELD_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --fieldId 330f29c5-5c4c-465f-9f4b-7903020ae1ce --listId 1f187321-f086-4d3d-8523-517e94cc9df9 --viewId 3d760127-982c-405e-9c93-e1f76e1a1110

    Add field with title ${chalk.grey('Custom field')} to view with title ${chalk.grey('Custom view')} from the list with title ${chalk.grey('Documents')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_VIEW_FIELD_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --fieldTitle 'Custom field' --listTitle Documents --viewTitle 'Custom view'
    
    Add field with title ${chalk.grey('Custom field')} at the position 0 to view with title ${chalk.grey('Custom view')} from the list with title ${chalk.grey('Documents')} located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_VIEW_FIELD_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --fieldTitle 'Custom field' --listTitle Documents --viewTitle 'Custom view' --fieldPosition 0
      `);
  }
}

module.exports = new SpoListViewFieldAddCommand();