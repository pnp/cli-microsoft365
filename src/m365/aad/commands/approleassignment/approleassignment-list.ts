import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import AadCommand from '../../../base/AadCommand';
import request from '../../../../request';
import { AppRoleAssignment } from './AppRoleAssignment';
import { ServicePrincipal, AppRole } from './ServicePrincipal';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  displayName?: string;
  objectId?: string;
}

class AadAppRoleAssignmentListCommand extends AadCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Lists app role assignments for the specified application registration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.displayName = typeof args.options.displayName !== 'undefined';
    telemetryProps.objectId = typeof args.options.objectId !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let sp: ServicePrincipal;

    // get the service principal associated with the appId
    let spMatchQuery: string = '';
    if (args.options.appId) {
      spMatchQuery = `appId eq '${encodeURIComponent(args.options.appId)}'`;
    }
    else if (args.options.objectId) {
      spMatchQuery = `objectId eq '${encodeURIComponent(args.options.objectId)}'`;
    }
    else {
      spMatchQuery = `displayName eq '${encodeURIComponent(args.options.displayName as string)}'`;
    }

    this
      .getServicePrincipalForApp(spMatchQuery)
      .then((resp: { value: ServicePrincipal[] }): Promise<ServicePrincipal[]> => {
        if (!resp.value.length) {
          return Promise.reject('app registration not found');
        }

        sp = resp.value[0];

        // the role assignment has an appRoleId but no name. To get the name,
        // we need to get all the roles from the resource. the resource is
        // a service principal. Multiple roles may have same resource id.
        const resourceIds = sp.appRoleAssignments.map((item: AppRoleAssignment) => item.resourceId);

        const tasks: Promise<ServicePrincipal>[] = [];
        for (let i: number = 0; i < resourceIds.length; i++) {
          tasks.push(this.getServicePrincipal(resourceIds[i]));
        }

        return Promise.all(tasks);
      })
      .then((resources: ServicePrincipal[]) => {
        // loop through all appRoleAssignments for the servicePrincipal
        // and lookup the appRole.Id in the resources[resourceId].appRoles array...
        const results: any[] = [];
        sp.appRoleAssignments.map((appRoleAssignment: AppRoleAssignment) => {
          const resource: ServicePrincipal | undefined = resources.find((r: any) => r.objectId === appRoleAssignment.resourceId);

          if (resource) {
            const appRole: AppRole | undefined = resource.appRoles.find((r: any) => r.id === appRoleAssignment.id);

            if (appRole) {
              results.push({
                appRoleId: appRoleAssignment.id,
                resourceDisplayName: appRoleAssignment.resourceDisplayName,
                resourceId: appRoleAssignment.resourceId,
                roleId: appRole.id,
                roleName: appRole.value
              });
            }
          }
        });

        if (args.options.output === 'json') {
          cmd.log(results);
        }
        else {
          cmd.log(results.map((r: any) => {
            return {
              resourceDisplayName: r.resourceDisplayName,
              roleName: r.roleName
            }
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getServicePrincipalForApp(filterParam: string): Promise<{ value: ServicePrincipal[] }> {
    const spRequestOptions: any = {
      url: `${this.resource}/myorganization/servicePrincipals?api-version=1.6&$expand=appRoleAssignments&$filter=${filterParam}`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    return request.get<{ value: ServicePrincipal[] }>(spRequestOptions);
  }

  private getServicePrincipal(spId: string): Promise<ServicePrincipal> {
    const spRequestOptions: any = {
      url: `${this.resource}/myorganization/servicePrincipals/${spId}?api-version=1.6`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    return request.get<ServicePrincipal>(spRequestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --appId [appId]',
        description: 'Application (client) Id of the App Registration for which the configured app roles should be retrieved'
      },
      {
        option: '-n, --displayName [displayName]',
        description: 'Display name of the application for which the configured app roles should be retrieved'
      },
      {
        option: '--objectId [objectId]',
        description: 'ObjectId of the application for which the configured app roles should be retrieved'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.appId && !args.options.displayName && !args.options.objectId) {
        return 'Specify either appId, objectId or displayName';
      }

      if (args.options.appId && !Utils.isValidGuid(args.options.appId)) {
        return `${args.options.appId} is not a valid GUID`;
      }

      if (args.options.objectId && !Utils.isValidGuid(args.options.objectId)) {
        return `${args.options.objectId} is not a valid GUID`;
      }

      let optionsSpecified: number = 0;
      optionsSpecified += args.options.appId ? 1 : 0;
      optionsSpecified += args.options.displayName ? 1 : 0;
      optionsSpecified += args.options.objectId ? 1 : 0;
      if (optionsSpecified > 1) {
        return 'Specify either appId, objectId or displayName';
      }

      return true;
    };
  }
}

module.exports = new AadAppRoleAssignmentListCommand();