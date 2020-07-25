import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import AadCommand from '../../../base/AadCommand';
import Utils from '../../../../Utils';
import { ServicePrincipal } from './ServicePrincipal';
import * as os from 'os';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface AppRole {
  objectId: string;
  value: string;
  resourceId: string;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  objectId?: string;
  displayName?: string;
  resource: string;
  scope: string;
}

class AadAppRoleAssignmentAddCommand extends AadCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds service principal permissions also known as scopes and app role assignments for specified Azure AD application registration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.objectId = typeof args.options.objectId !== 'undefined';
    telemetryProps.displayName = typeof args.options.displayName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let objectId: string = '';
    let queryFilter: string = '';
    if (args.options.appId) {
      queryFilter = `$filter=appId eq '${encodeURIComponent(args.options.appId)}'`;
    }
    else if (args.options.objectId) {
      queryFilter = `$filter=objectId eq '${encodeURIComponent(args.options.objectId)}'`;
    }
    else {
      queryFilter = `$filter=displayName eq '${encodeURIComponent(args.options.displayName as string)}'`;
    }

    const getServicePrinciplesRequestOptions: any = {
      url: `${this.resource}/myorganization/servicePrincipals?api-version=1.6&${queryFilter}`,
      headers: {
        accept: 'application/json;odata=nometadata;streaming=false'
      },
      json: true
    };

    request
      .get<{ value: ServicePrincipal[] }>(getServicePrinciplesRequestOptions)
      .then((servicePrincipalResult: { value: ServicePrincipal[] }): Promise<{ value: ServicePrincipal[] }> => {
        if (servicePrincipalResult.value.length > 1) {
          return Promise.reject('More than one service principal found. Please use the appId or objectId option to make sure the right service principal is specified.');
        }

        objectId = servicePrincipalResult.value[0].objectId;

        let resource: string = encodeURIComponent(args.options.resource);

        // try resolve aliases that the user might enter since these are seen in the Azure portal
        switch (args.options.resource.toLocaleLowerCase()) {
          case 'sharepoint':
            resource = 'Microsoft 365 SharePoint Online';
            break;
          case 'intune':
            resource = 'Microsoft Intune API';
            break;
          case 'exchange':
            resource = 'Microsoft 365 Exchange Online';
            break;
        }

        // will perform resource name, appId or objectId search
        let filter: string = `$filter=publisherName eq '${resource}' or (displayName eq '${resource}' or startswith(displayName,'${resource}'))`;

        if (Utils.isValidGuid(resource)) {
          filter += ` or appId eq '${resource}' or objectId eq '${resource}'`;
        }

        const requestOptions: any = {
          url: `${this.resource}/myorganization/servicePrincipals?api-version=1.6&${filter}`,
          headers: {
            'accept': 'application/json;odata=nometadata;streaming=false'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res: { value: ServicePrincipal[] }): Promise<AppRole[]> => {
        const result: AppRole[] = [];

        // flatten the app roles found
        const appRolesFound: AppRole[] = [];
        for (const servicePrincipal of res.value) {
          for (const role of servicePrincipal.appRoles) {
            appRolesFound.push({
              resourceId: servicePrincipal.objectId,
              objectId: role.id,
              value: role.value
            });
          }
        }

        if (!appRolesFound.length) {
          return Promise.reject(`The resource '${args.options.resource}' does not have any application permissions available.`);
        }

        // search for match between the found app roles and the specified scope option value
        for (const scope of args.options.scope.split(',')) {
          const existingRoles = appRolesFound.filter((role: AppRole) => {
            return role.value.toLocaleLowerCase() === scope.toLocaleLowerCase().trim();
          });

          if (!existingRoles.length) {
            // the role specified in the scope option does not belong to the found service principles
            // throw an error and show list with available roles (scopes)
            let availableRoles: string = '';
            appRolesFound.map((r: AppRole) => availableRoles += `${os.EOL}${r.value}`);

            return Promise.reject(`The scope value '${scope}' you have specified does not exist for ${args.options.resource}. ${os.EOL}Available scopes (application permissions) are: ${availableRoles}`);
          }

          result.push(existingRoles[0]);
        }

        return Promise.resolve(result);
      })
      .then((appRoles: AppRole[]) => {
        const tasks: Promise<any>[] = [];

        for (const appRole of appRoles) {
          tasks.push(this.addRoleToServicePrincipal(objectId, appRole));
        }

        return Promise.all(tasks);
      })
      .then((rolesAddedResponse: any) => {
        if (args.options.output && args.options.output.toLowerCase() === 'json') {
          cmd.log(rolesAddedResponse);
        }
        else {
          cmd.log(rolesAddedResponse.map((result: any) => ({
            objectId: result.objectId,
            principalDisplayName: result.principalDisplayName,
            resourceDisplayName: result.resourceDisplayName
          })));
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  private addRoleToServicePrincipal(objectId: string, appRole: AppRole): Promise<any> {
    const requestOptions: any = {
      url: `${this.resource}/myorganization/servicePrincipals/${objectId}/appRoleAssignments?api-version=1.6`,
      headers: {
        'accept': 'application/json;odata=nometadata;streaming=false',
        'Content-Type': 'application/json'
      },
      json: true,
      body: {
        id: appRole.objectId,
        principalId: objectId,
        resourceId: appRole.resourceId
      }
    };

    return request.post(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--appId [appId]',
        description: 'Application appId also known as clientId of the App Registration to which the configured scopes (app roles) should be applied'
      },
      {
        option: '--objectId [objectId]',
        description: 'Application objectId of the App Registration to which the configured scopes (app roles) should be applied'
      },
      {
        option: '--displayName [displayName]',
        description: 'Application name of the App Registration to which the configured scopes (app roles) should be applied'
      },
      {
        option: '-r, --resource <resource>',
        description: 'Service principal name, appId or objectId that has the scopes (roles) ex. SharePoint',
        autocomplete: ['Microsoft Graph', 'SharePoint', 'OneNote', 'Exchange', 'Microsoft Forms', 'Azure Active Directory Graph', 'Skype for Business']
      },
      {
        option: '-s, --scope <scope>',
        description: 'Permissions known also as scopes and roles to grant the application with. If multiple permissions have to be granted, they have to be comma separated ex. \'Sites.Read.All,Sites.ReadWrite.all\''
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      let optionsSpecified: number = 0;
      optionsSpecified += args.options.appId ? 1 : 0;
      optionsSpecified += args.options.displayName ? 1 : 0;
      optionsSpecified += args.options.objectId ? 1 : 0;
      if (optionsSpecified !== 1) {
        return 'Specify either appId, objectId or displayName';
      }

      if (args.options.appId && !Utils.isValidGuid(args.options.appId)) {
        return `${args.options.appId} is not a valid GUID`;
      }

      if (args.options.objectId && !Utils.isValidGuid(args.options.objectId)) {
        return `${args.options.objectId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new AadAppRoleAssignmentAddCommand();