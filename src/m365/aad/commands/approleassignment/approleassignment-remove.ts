import * as os from 'os';
import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { AppRoleAssignment } from './AppRoleAssignment';
import { AppRole, ServicePrincipal } from './ServicePrincipal';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appDisplayName?: string;
  appObjectId?: string;
  resource: string;
  scope: string;
  confirm?: boolean;
}

class AadAppRoleAssignmentRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Deletes an app role assignment for the specified Azure AD Application Registration';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        appId: typeof args.options.appId !== 'undefined',
        appDisplayName: typeof args.options.appDisplayName !== 'undefined',
        appObjectId: typeof args.options.appObjectId !== 'undefined',
        confirm: (!!args.options.confirm).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--appId [appId]'
      },
      {
        option: '--appObjectId [appObjectId]'
      },
      {
        option: '--appDisplayName [appDisplayName]'
      },
      {
        option: '-r, --resource <resource>',
        autocomplete: ['Microsoft Graph', 'SharePoint', 'OneNote', 'Exchange', 'Microsoft Forms', 'Azure Active Directory Graph', 'Skype for Business']
      },
      {
        option: '-s, --scope <scope>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        let optionsSpecified: number = 0;
        optionsSpecified += args.options.appId ? 1 : 0;
        optionsSpecified += args.options.appDisplayName ? 1 : 0;
        optionsSpecified += args.options.appObjectId ? 1 : 0;
        if (optionsSpecified !== 1) {
          return 'Specify either appId, appObjectId or appDisplayName';
        }
    
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }
    
        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
          return `${args.options.appObjectId} is not a valid GUID`;
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeAppRoleAssignment: () => void = (): void => {
      let sp: ServicePrincipal;
      // get the service principal associated with the appId
      let spMatchQuery: string = '';
      if (args.options.appId) {
        spMatchQuery = `appId eq '${encodeURIComponent(args.options.appId)}'`;
      }
      else if (args.options.appObjectId) {
        spMatchQuery = `id eq '${encodeURIComponent(args.options.appObjectId)}'`;
      }
      else {
        spMatchQuery = `displayName eq '${encodeURIComponent(args.options.appDisplayName as string)}'`;
      }

      this
        .getServicePrincipalForApp(spMatchQuery)
        .then((resp: { value: ServicePrincipal[] }): Promise<{ value: ServicePrincipal[] }> => {
          if (!resp.value.length) {
            return Promise.reject('app registration not found');
          }

          sp = resp.value[0];
          let resource: string = encodeURIComponent(args.options.resource);

          // try resolve aliases that the user might enter since these are seen in the Azure portal
          switch (args.options.resource.toLocaleLowerCase()) {
            case 'sharepoint':
              resource = 'Office 365 SharePoint Online';
              break;
            case 'intune':
              resource = 'Microsoft Intune API';
              break;
            case 'exchange':
              resource = 'Office 365 Exchange Online';
              break;
          }

          // will perform resource name, appId or objectId search
          let filter: string = `$filter=(displayName eq '${resource}' or startswith(displayName,'${resource}'))`;

          if (validation.isValidGuid(resource)) {
            filter += ` or appId eq '${resource}' or id eq '${resource}'`;
          }
          const requestOptions: any = {
            url: `${this.resource}/v1.0/servicePrincipals?${filter}`,
            headers: {
              'accept': 'application/json'
            },
            responseType: 'json'
          };

          return request.get<{ value: ServicePrincipal[] }>(requestOptions);
        })
        .then((resp: { value: ServicePrincipal[] }): Promise<AppRole[]> => {
          if (!resp.value.length) {
            return Promise.reject('Resource not found');
          }

          const appRolesToBeDeleted: AppRole[] = [];
          const appRolesFound: AppRole[] = resp.value[0].appRoles;

          if (!appRolesFound.length) {
            return Promise.reject(`The resource '${args.options.resource}' does not have any application permissions available.`);
          }

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

            appRolesToBeDeleted.push(existingRoles[0]);
          }
          const tasks: Promise<any>[] = [];

          for (const appRole of appRolesToBeDeleted) {
            const appRoleAssignment = sp.appRoleAssignments.filter((role: AppRoleAssignment) => role.appRoleId === appRole.id);
            if (!appRoleAssignment.length) {
              return Promise.reject('App role assignment not found');
            }
            tasks.push(this.removeAppRoleAssignmentForServicePrincipal(sp.id, appRoleAssignment[0].id));
          }

          return Promise.all(tasks);
        }).then(_ => cb(), (res: any): void => this.handleRejectedODataJsonPromise(res, logger, cb));
    };

    if (args.options.confirm) {
      removeAppRoleAssignment();
    }
    else {
      Cli.prompt(
        {
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove the appRoleAssignment with scope ${args.options.scope} for resource ${args.options.resource}?`
        },
        (result: { continue: boolean }): void => {
          if (!result.continue) {
            cb();
          }
          else {
            removeAppRoleAssignment();
          }
        }
      );
    }
  }

  private getServicePrincipalForApp(filterParam: string): Promise<{ value: ServicePrincipal[] }> {
    const spRequestOptions: any = {
      url: `${this.resource}/v1.0/servicePrincipals?$expand=appRoleAssignments&$filter=${filterParam}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request.get<{ value: ServicePrincipal[] }>(spRequestOptions);
  }

  private removeAppRoleAssignmentForServicePrincipal(spId: string, appRoleAssignmentId: string): Promise<ServicePrincipal> {
    const spRequestOptions: any = {
      url: `${this.resource}/v1.0/servicePrincipals/${spId}/appRoleAssignments/${appRoleAssignmentId}`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.delete(spRequestOptions);
  }
}

module.exports = new AadAppRoleAssignmentRemoveCommand();