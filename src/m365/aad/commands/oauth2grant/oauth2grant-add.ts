import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import AadCommand from '../../../base/AadCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  clientId: string;
  resourceId: string;
  scope: string;
}

class AadOAuth2GrantAddCommand extends AadCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_ADD;
  }

  public get description(): string {
    return 'Grant the specified service principal OAuth2 permissions to the specified resource';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Granting the service principal specified permissions...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/myorganization/oauth2PermissionGrants?api-version=1.6`,
      headers: {
        'content-type': 'application/json'
      },
      json: true,
      body: {
        "odata.type": "Microsoft.DirectoryServices.OAuth2PermissionGrant",
        "clientId": args.options.clientId,
        "consentType": "AllPrincipals",
        "principalId": null,
        "resourceId": args.options.resourceId,
        "scope": args.options.scope,
        "startTime": "0001-01-01T00:00:00",
        "expiryTime": "9000-01-01T00:00:00"
      }
    };

    request
      .post<void>(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --clientId <clientId>',
        description: 'objectId of the service principal for which permissions should be granted'
      },
      {
        option: '-r, --resourceId <resourceId>',
        description: 'objectId of the AAD application to which permissions should be granted'
      },
      {
        option: '-s, --scope <scope>',
        description: 'Permissions to grant'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.clientId) {
        return 'Required option clientId missing';
      }

      if (!Utils.isValidGuid(args.options.clientId)) {
        return `${args.options.clientId} is not a valid GUID`;
      }

      if (!args.options.resourceId) {
        return 'Required option resourceId missing';
      }

      if (!Utils.isValidGuid(args.options.resourceId)) {
        return `${args.options.resourceId} is not a valid GUID`;
      }

      if (!args.options.scope) {
        return 'Required option scope missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.OAUTH2GRANT_ADD).helpInformation());
    log(
      `  Remarks:
  
    Before you can grant service principal OAuth2 permissions, you need its ${chalk.grey('objectId')}.
    You can retrieve it using the ${chalk.blue(commands.SP_GET)} command.

    The resource for which you want to grant permissions is designated using its ${chalk.grey('objectId')}.
    You can retrieve it using the ${chalk.blue(commands.SP_GET)} command, the same way you would retrieve
    the ${chalk.grey('objectId')} of the service principal.

    When granting OAuth2 permissions, you have to specify which permission scopes you want to grant
    the service principal. You can get the list of available permission scopes either from the resource
    documentation or from the ${chalk.grey('appRoles')} property when retrieving information
    about the service principal using the ${chalk.blue(commands.SP_GET)} command. Multiple permission
    scopes can be specified separated by a space.

    When granting OAuth2 permissions, the values of the ${chalk.grey('clientId')} and ${chalk.grey('resourceId')}
    properties form a unique key. If a grant for the same ${chalk.grey('clientId')}-${chalk.grey('resourceId')}
    pair already exists, running the ${chalk.blue(commands.OAUTH2GRANT_ADD)} command will fail with an error.
    If you want to change permissions on an existing OAuth2 grant use the ${chalk.blue(commands.OAUTH2GRANT_SET)}
    command instead.
   
  Examples:
  
    Grant the service principal ${chalk.grey('d03a0062-1aa6-43e1-8f49-d73e969c5812')} the
    ${chalk.grey('Calendars.Read')} OAuth2 permissions to the ${chalk.grey('c2af2474-2c95-423a-b0e5-e4895f22f9e9')} resource.
      ${commands.OAUTH2GRANT_ADD} --clientId d03a0062-1aa6-43e1-8f49-d73e969c5812 --resourceId c2af2474-2c95-423a-b0e5-e4895f22f9e9 --scope Calendars.Read

    Grant the service principal ${chalk.grey('d03a0062-1aa6-43e1-8f49-d73e969c5812')} the
    ${chalk.grey('Calendars.Read')} and ${chalk.grey('Mail.Read')} OAuth2 permissions to the ${chalk.grey('c2af2474-2c95-423a-b0e5-e4895f22f9e9')} resource.
      ${commands.OAUTH2GRANT_ADD} --clientId d03a0062-1aa6-43e1-8f49-d73e969c5812 --resourceId c2af2474-2c95-423a-b0e5-e4895f22f9e9 --scope "Calendars.Read Mail.Read"

  More information:
  
    Application and service principal objects in Azure Active Directory (Azure AD)
      https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects
`);
  }
}

module.exports = new AadOAuth2GrantAddCommand();