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
import { ServicePrincipal } from './ServicePrincipal';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  displayName?: string;
}

class AadAppRoleAssignmentListCommand extends AadCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Lists AppRoleAssignments for the specified application registration';
  }

  public async commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): Promise<void> {
    try {
      // get the service principal associated with the appId
      const spMatchQuery: string = args.options.appId ?
        `appId eq '${encodeURIComponent(args.options.appId)}'` :
        `displayName eq '${encodeURIComponent(args.options.displayName as string)}'`;

      let sp: ServicePrincipal = await this.GetServicePrincipalForApp(spMatchQuery);

      if (!sp) {
        cmd.log('app registration not found');
        cb();
        return;
      }

      // The role assignment has an appRoleId but no name. To get the name, we need to get all the roles from the resource.
      // The resource is a service principal. Multiple roles may have same resource id.
      let resourceIds = Array.from(new Set(sp.appRoleAssignments.map((item: AppRoleAssignment) => item.resourceId)));

      let resources: ServicePrincipal[] = [];

      for (var i = 0; i < resourceIds.length; i++) {
        let resource = await this.GetServicePrincipal(resourceIds[i]);
        resources.push(resource);
      }

      // resourceIds.forEach(async id => {
      //   let resource = await this.GetServicePrincipal(id);
      //   resources.push(resource);
      // });

      // loop thru all appRoleAssignments for the servicePrincipal and lookup the appRole.Id in the resources[resourceId].appRoles array...
      let results: any[] = [];
      sp.appRoleAssignments.map((appRoleAssignment: AppRoleAssignment) => {
        let resource = resources.find(r => r.objectId === appRoleAssignment.resourceId);
        if (resource) {
          let appRole = resource.appRoles.find(r => r.id === appRoleAssignment.id)
          if (appRole) {
            results.push({
              appRoleId: appRoleAssignment.id,
              resourceDisplayName: appRoleAssignment.resourceDisplayName,
              resourceId: appRoleAssignment.resourceId,
              roleId: appRole.id,
              roleName: appRole.value
            })
          }
        }
      });



      if (args.options.output === 'json') {
        cmd.log(results);
      }
      else {
        cmd.log(results.map(r => {
          return {
            resourceDisplayName: r.resourceDisplayName,
            roleName: r.roleName
          }
        }));
      }

      cb();
    } catch (error) {
      this.handleRejectedODataJsonPromise(error, cmd, cb)
    }

  }

  private async GetServicePrincipalForApp(filterParam: string): Promise<ServicePrincipal> {

    const spRequestOptions: any = {
      url: `${this.resource}/myorganization/servicePrincipals?api-version=1.6&$expand=appRoleAssignments&$filter=${filterParam}`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    let sp = await request.get<{ value: ServicePrincipal[] }>(spRequestOptions)
      .then((response: { value: ServicePrincipal[] }): ServicePrincipal => {
        return response.value[0];
      });

    return sp;
  }

  private async GetServicePrincipal(spId: string): Promise<ServicePrincipal> {
    const spRequestOptions: any = {
      url: `${this.resource}/myorganization/servicePrincipals/${spId}?api-version=1.6`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    let sp = await request.get<ServicePrincipal>(spRequestOptions)
      .then((response) => {
        return response;
      });

    return sp;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --appId [appId]',
        description: 'Application (client) Id of the App Registration for which the configured appRoles should be retrieved'
      },
      {
        option: '-n, --displayName [displayName]',
        description: 'Display name of the application for which the configured appRoles should be retrieved'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.appId && !args.options.displayName) {
        return 'Specify either appId or displayName';
      }

      if (args.options.appId) {
        if (!Utils.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }
      }

      if (args.options.appId && args.options.displayName) {
        return 'Specify either appId or displayName but not both';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APPROLEASSIGNMENT_LIST).helpInformation());
    log(
      `  Remarks:
  
    Specify either the ${chalk.grey('appId')} or ${chalk.grey('displayName')} but not both. If you specify both values, the command will fail
    with an error.
   
  Examples:
  
    List AppRoles assigned to service principal with Application (client) ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}.
      ${commands.APPROLEASSIGNMENT_LIST} --appId b2307a39-e878-458b-bc90-03bc578531d6

  More information:
  
  Application and service principal objects in Azure Active Directory (Azure AD): 
  https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects
`);
  }
}

module.exports = new AadAppRoleAssignmentListCommand();