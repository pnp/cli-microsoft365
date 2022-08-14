import * as os from 'os';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { ServicePrincipal } from './ServicePrincipal';

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

class AadAppRoleAssignmentAddCommand extends GraphCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds service principal permissions also known as scopes and app role assignments for specified Azure AD application registration';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        appId: typeof args.options.appId !== 'undefined',
        objectId: typeof args.options.objectId !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--appId [appId]'
      },
      {
        option: '--objectId [objectId]'
      },
      {
        option: '--displayName [displayName]'
      },
      {
        option: '-r, --resource <resource>',
        autocomplete: ['Microsoft Graph', 'SharePoint', 'OneNote', 'Exchange', 'Microsoft Forms', 'Azure Active Directory Graph', 'Skype for Business']
      },
      {
        option: '-s, --scope <scope>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
          return `${args.options.objectId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['appId', 'objectId', 'displayName']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let objectId: string = '';
    let queryFilter: string = '';
    if (args.options.appId) {
      queryFilter = `$filter=appId eq '${encodeURIComponent(args.options.appId)}'`;
    }
    else if (args.options.objectId) {
      queryFilter = `$filter=id eq '${encodeURIComponent(args.options.objectId)}'`;
    }
    else {
      queryFilter = `$filter=displayName eq '${encodeURIComponent(args.options.displayName as string)}'`;
    }

    const getServicePrinciplesRequestOptions: any = {
      url: `${this.resource}/v1.0/servicePrincipals?${queryFilter}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .get<{ value: ServicePrincipal[] }>(getServicePrinciplesRequestOptions)
      .then((servicePrincipalResult: { value: ServicePrincipal[] }): Promise<{ value: ServicePrincipal[] }> => {
        if (servicePrincipalResult.value.length > 1) {
          return Promise.reject('More than one service principal found. Please use the appId or objectId option to make sure the right service principal is specified.');
        }

        objectId = servicePrincipalResult.value[0].id;

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

        return request.get(requestOptions);
      })
      .then((res: { value: ServicePrincipal[] }): Promise<AppRole[]> => {
        const result: AppRole[] = [];

        // flatten the app roles found
        const appRolesFound: AppRole[] = [];
        for (const servicePrincipal of res.value) {
          for (const role of servicePrincipal.appRoles) {
            appRolesFound.push({
              resourceId: servicePrincipal.id,
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
          logger.log(rolesAddedResponse);
        }
        else {
          logger.log(rolesAddedResponse.map((result: any) => ({
            objectId: result.id,
            principalDisplayName: result.principalDisplayName,
            resourceDisplayName: result.resourceDisplayName
          })));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private addRoleToServicePrincipal(objectId: string, appRole: AppRole): Promise<any> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/servicePrincipals/${objectId}/appRoleAssignments`,
      headers: {
        'Content-Type': 'application/json'
      },
      responseType: 'json',
      data: {
        appRoleId: appRole.objectId,
        principalId: objectId,
        resourceId: appRole.resourceId
      }
    };

    return request.post(requestOptions);
  }
}

module.exports = new AadAppRoleAssignmentAddCommand();