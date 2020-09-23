import { Logger } from '../../../../cli';
import { CommandError, CommandOption } from '../../../../Command';
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
  url: string;
}

class SpoSiteAppCatalogRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_APPCATALOG_REMOVE;
  }

  public get description(): string {
    return 'Removes site collection app catalog from the specified site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.url = (!(!args.options.url)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const url: string = args.options.url;
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          logger.log(`Disabling site collection app catalog...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="50" ObjectPathId="49" /><ObjectPath Id="52" ObjectPathId="51" /><ObjectPath Id="54" ObjectPathId="53" /><ObjectPath Id="56" ObjectPathId="55" /><ObjectPath Id="58" ObjectPathId="57" /><Method Name="Remove" Id="59" ObjectPathId="57"><Parameters><Parameter Type="String">${Utils.escapeXml(url)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="49" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="51" ParentId="49" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(url)}</Parameter></Parameters></Method><Property Id="53" ParentId="51" Name="RootWeb" /><Property Id="55" ParentId="53" Name="TenantAppCatalog" /><Property Id="57" ParentId="55" Name="SiteCollectionAppCatalogsSites" /></ObjectPaths></Request>`
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
            logger.log('Site collection app catalog disabled');
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the site collection containing the app catalog to disable'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.url);
  }
}

module.exports = new SpoSiteAppCatalogRemoveCommand();