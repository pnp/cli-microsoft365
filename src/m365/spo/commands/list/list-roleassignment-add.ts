import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, urlUtil, validation } from '../../../../utils';
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
  principalId?: number;
  upn?: string;
  groupName?: string;
  roleDefinitionId?: string;
  roleDefinitionName?: string;
}

class SpoListRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to list permissions';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.listUrl = typeof args.options.listUrl !== 'undefined';
    telemetryProps.principalId = typeof args.options.principalId !== 'undefined';
    telemetryProps.upn = typeof args.options.upn !== 'undefined';
    telemetryProps.groupName = typeof args.options.groupName !== 'undefined';
    telemetryProps.roleDefinitionId = typeof args.options.roleDefinitionId !== 'undefined';
    telemetryProps.roleDefinitionName = typeof args.options.roleDefinitionName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Adding role assignment to list in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/`;
    if (args.options.listId) {
      requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
    }
    else if (args.options.listTitle) {
      requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
    }

    if (args.options.upn) {
      // TODO: Add support for adding role assignment to user by UPN
    } 
    else if (args.options.groupName) {
      // TODO: Add support for adding role assignment to group by name
    }

    if (args.options.roleDefinitionName) {
      // TODO: Add support for adding role assignment by role definition name
    }

    const requestOptions: any = {
      url: `${requestUrl}/roleassignments/addroleassignment(principalid='${args.options.principalId}',roledefid='${args.options.roleDefinitionId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--principalId [principalId]'
      },
      {
        option: '--upn [upn]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--roleDefinitionId [roleDefinitionId]'
      },
      {
        option: '--roleDefinitionName [roleDefinitionName]'
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

    if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
      return `${args.options.listId} is not a valid GUID`;
    }

    if (args.options.principalId && isNaN(args.options.principalId)) {
      return `Specified principalId ${args.options.principalId} is not a number`;
    }

    const listOptions: any[] = [args.options.listId, args.options.listTitle, args.options.listUrl];
    if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
      return `Specify either list id or title or list url`;
    }

    const principalOptions: any[] = [args.options.principalId, args.options.upn, args.options.groupName];
    if (principalOptions.some(item => item !== undefined) && principalOptions.filter(item => item !== undefined).length > 1) {
      return `Specify either principalId id or upn or groupName`;
    }

    const roleDefinitionOptions: any[] = [args.options.roleDefinitionId, args.options.roleDefinitionName];
    if (roleDefinitionOptions.some(item => item !== undefined) && roleDefinitionOptions.filter(item => item !== undefined).length > 1) {
      return `Specify either roleDefinitionId id or roleDefinitionName`;
    }

    return true;
  }
}

module.exports = new SpoListRoleAssignmentAddCommand();