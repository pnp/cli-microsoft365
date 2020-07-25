import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  enabled: string;
  confirm?: boolean;
}

class SpoServicePrincipalSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SERVICEPRINCIPAL_SET;
  }

  public get description(): string {
    return 'Enable or disable the service principal';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.enabled = args.options.enabled === 'true';
    return telemetryProps;
  }

  public alias(): string[] | undefined {
    return [commands.SP_SET];
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const enabled: boolean = args.options.enabled === 'true';

    const toggleServicePrincipal: () => void = (): void => {
      let spoAdminUrl: string = '';

      this
        .getSpoAdminUrl(cmd, this.debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;

          return this.getRequestDigest(spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            cmd.log(`${(enabled ? 'Enabling' : 'Disabling')} service principal...`);
          }

          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><SetProperty Id="29" ObjectPathId="27" Name="AccountEnabled"><Parameter Type="Boolean">${enabled}</Parameter></SetProperty><Method Name="Update" Id="30" ObjectPathId="27" /><Query Id="31" ObjectPathId="27"><Query SelectAllProperties="true"><Properties><Property Name="AccountEnabled" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="27" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            cb(new CommandError(response.ErrorInfo.ErrorMessage));
            return;
          }
          else {
            const output: any = json[json.length - 1];
            delete output._ObjectType_;

            cmd.log(output);

            if (this.verbose) {
              cmd.log(chalk.green('DONE'));
            }
          }
          cb();
        }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
    }

    if (args.options.confirm) {
      toggleServicePrincipal();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to ${enabled ? 'enable' : 'disable'} the service principal?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          toggleServicePrincipal();
        }
      });
    }
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const enabled: string = args.options.enabled.toLowerCase();
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
        description: 'Set to true to enable the service principal or to false to disable it. Valid values are true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming enabling/disabling the service principal'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoServicePrincipalSetCommand();