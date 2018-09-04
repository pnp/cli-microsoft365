import auth from '../../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type: string;
  enabled: string;
  noDefaultOrigins?: boolean;
}

class SpoCdnSetCommand extends SpoCommand {
  public get name(): string {
    return commands.CDN_SET;
  }

  public get description(): string {
    return 'Enable or disable the specified Office 365 CDN';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.cdnType = args.options.type || 'Public';
    telemetryProps.enabled = args.options.enabled === 'true';
    telemetryProps.noDefaultOrigins = (!(!args.options.noDefaultOrigins)).toString();
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const cdnTypeString: string = args.options.type || 'Public';
    const enabled: boolean = args.options.enabled === 'true';
    let cdnType: number = 0;

    switch (cdnTypeString) {
      case "Private": {
        cdnType = 1;
        break;
      }
      case "Both": {
        cdnType = 2;
        break;
      }
      default: {
        cdnType = 0;
        break;
      }
    }

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Loading CDN settings for the ${auth.site.url} tenant...`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        let requestBody = '';

        if (cdnType === 2) {
          if (args.options.noDefaultOrigins) {
            if (this.verbose) {
              cmd.log(`${(enabled ? 'Enabling' : 'Disabling')} Both (Privated and Public) CDN locations without default originins. Please wait, this might take a moment...`);
            }

            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><Method Name="SetTenantCdnEnabled" Id="13" ObjectPathId="11"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="14" ObjectPathId="11"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
          }
          else {
            if (this.verbose) {
              cmd.log(`${(enabled ? 'Enabling' : 'Disabling')} Both (Privated and Public) CDN locations without default originins. Please wait, this might take a moment...`);
            }

            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="96" ObjectPathId="95" /><Method Name="SetTenantCdnEnabled" Id="97" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="98" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method>${(enabled ? '<Method Name="CreateTenantCdnDefaultOrigins" Id="99" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="100" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method>' : '')}</Actions><ObjectPaths><Constructor Id="95" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
          }
        }
        else {
          if (args.options.noDefaultOrigins) {
            if (this.verbose) {
              cmd.log(`${(enabled ? 'Enabling' : 'Disabling')} ${(cdnType === 1 ? 'Private' : 'Public')} CDN without default origins. Please wait, this might take a moment...`);
            }

            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
          }
          else {
            if (this.verbose) {
              cmd.log(`${(enabled ? 'Enabling' : 'Disabling')} ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
            }

            if (enabled) {
              requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">${(cdnType === 1 ? 1 : 0)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
            }
            else {
              requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
            }
          }
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: requestBody
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }
        else {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.type) {
        if (args.options.type !== 'Public' && args.options.type !== 'Both' &&
          args.options.type !== 'Private') {
          return `${args.options.type} is not a valid CDN type. Allowed values are Public|Private|Both`;
        }
      }

      const enabled: string | undefined = args.options.enabled ? args.options.enabled.toLowerCase() : undefined;
      if (enabled !== 'true' &&
        enabled !== 'false') {
        return `${args.options.enabled} is not a valid boolean value. Allowed values are true|false`;
      }

      return true;
    };
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-e, --enabled <enabled>',
        description: 'Set to true to enable CDN or to false to disable it. Valid values are true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '-t, --type [type]',
        description: 'Type of CDN to manage. Public|Private. Default Public',
        autocomplete: ['Public', 'Private', 'Both']
      },
      {
        option: '--noDefaultOrigins',
        description: 'Disable creation of default origins'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.CDN_SET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To enable or disable an Office 365 CDN, you have to first connect to a tenant admin site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
    If you are connected to a different site and will try to manage tenant properties,
    you will get an error.

    Using the ${chalk.blue('-t, --type')} option you can choose whether you want to manage the settings of
    the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

    Using the ${chalk.blue('-e, --enabled')} option you can specify whether the given CDN type should be
    enabled or disabled. Use ${chalk.grey('true')} to enable the specified CDN and ${chalk.grey('false')} to
    disable it.

    Using the ${chalk.blue('-noDefaultOrigins')} option you can specify to skip the creation of the default origins.  

  Examples:
  
    Enable the Office 365 Public CDN on the current tenant
      ${chalk.grey(config.delimiter)} ${commands.CDN_SET} -t Public -e true

    Disable the Office 365 Public CDN on the current tenant
      ${chalk.grey(config.delimiter)} ${commands.CDN_SET} -t Public -e false

    Enable the Office 365 Private CDN on the current tenant
      ${chalk.grey(config.delimiter)} ${commands.CDN_SET} -t Private -e true
  
    Enable the Office 365 Private and Public CDN on the current tenant with default Origins 
      ${chalk.grey(config.delimiter)} ${commands.CDN_SET} -t Both -e true 

    Enable the Office 365 Private and Public CDN on the current tenant without the default Origins 
      ${chalk.grey(config.delimiter)} ${commands.CDN_SET} -t Both -e true --noDefaultOrigins

  More information:

    General availability of Office 365 CDN
      https://dev.office.com/blogs/general-availability-of-office-365-cdn
`);
  }
}

module.exports = new SpoCdnSetCommand();