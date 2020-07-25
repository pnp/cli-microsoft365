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
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  designPackageId?: string;
  url: string;
}

class SpoSiteCommSiteEnableCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITE_COMMSITE_ENABLE}`;
  }

  public get description(): string {
    return 'Enables communication site features on the specified site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.designPackageId = typeof args.options.designPackageId !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const designPackageId: string = args.options.designPackageId || '{d604dac3-50d3-405e-9ab9-d4713cda74ef}';
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        if (this.debug) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.url)}</Parameter><Parameter Type="Guid">${Utils.escapeXml(designPackageId)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The URL of the site to enable communication site features on'
      },
      {
        option: '-i, --designPackageId [designPackageId]',
        description: 'The ID of the site design to apply when enabling communication site features'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.designPackageId &&
        !Utils.isValidGuid(args.options.designPackageId)) {
        return `${args.options.designPackageId} is not a valid GUID`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.url);
    };
  }
}

module.exports = new SpoSiteCommSiteEnableCommand();