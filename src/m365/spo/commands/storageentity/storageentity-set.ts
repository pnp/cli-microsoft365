import { Logger } from '../../../../cli';
import {
  CommandError, CommandOption
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo } from '../../spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl: string;
  key: string;
  value: string;
  description?: string;
  comment?: string;
}

class SpoStorageEntitySetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.STORAGEENTITY_SET}`;
  }

  public get description(): string {
    return 'Sets tenant property on the specified SharePoint Online app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = (!(!args.options.description)).toString();
    telemetryProps.comment = (!(!args.options.comment)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          logger.logToStderr(`Setting tenant property ${args.options.key} in ${args.options.appCatalogUrl}...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.key)}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.value)}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.description || '')}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.comment || '')}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.appCatalogUrl)}</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          if (this.verbose && response.ErrorInfo.ErrorMessage.indexOf('Access denied.') > -1) {
            logger.logToStderr('');
            logger.logToStderr(`This error is often caused by invalid URL of the app catalog site. Verify, that the URL you specified as an argument of the ${commands.STORAGEENTITY_SET} command is a valid app catalog URL and try again.`);
            logger.logToStderr('');
          }

          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }
        
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --appCatalogUrl <appCatalogUrl>'
      },
      {
        option: '-k, --key <key>'
      },
      {
        option: '-v, --value <value>'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-c, --comment [comment]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const result: boolean | string = SpoCommand.isValidSharePointUrl(args.options.appCatalogUrl);
    if (result === false) {
      return 'Missing required option appCatalogUrl';
    }
    else {
      return result;
    }
  }
}

module.exports = new SpoStorageEntitySetCommand();