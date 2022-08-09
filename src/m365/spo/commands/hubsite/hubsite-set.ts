import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { HubSiteProperties } from './HubSiteProperties';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  title?: string;
  description?: string;
  logoUrl?: string;
}

class SpoHubSiteSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_SET;
  }

  public get description(): string {
    return 'Updates properties of the specified hub site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        title: typeof args.options.title === 'string',
        description: typeof args.options.description === 'string',
        logoUrl: typeof args.options.logoUrl === 'string'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-l, --logoUrl [logoUrl]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (!args.options.title &&
          !args.options.description &&
          !args.options.logoUrl) {
          return 'Specify title, description or logoUrl to update';
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('t', 'title', 'd', 'description', 'l', 'logoUrl');
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        return spo.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          logger.logToStderr(`Updating hub site ${args.options.id}...`);
        }

        const title: string = typeof args.options.title === 'string' ? `<SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">${formatting.escapeXml(args.options.title)}</Parameter></SetProperty>` : '';
        const description: string = typeof args.options.description === 'string' ? `<SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">${formatting.escapeXml(args.options.description)}</Parameter></SetProperty>` : '';
        const logoUrl: string = typeof args.options.logoUrl === 'string' ? `<SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">${formatting.escapeXml(args.options.logoUrl)}</Parameter></SetProperty>` : '';

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query>${title}${logoUrl}${description}<Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method></ObjectPaths></Request>`
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
          const hubSite: HubSiteProperties = json.pop();
          delete hubSite._ObjectType_;

          hubSite.ID = hubSite.ID.replace('/Guid(', '').replace(')/', '');
          hubSite.SiteId = hubSite.SiteId.replace('/Guid(', '').replace(')/', '');

          logger.log(hubSite);
        }

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
}

module.exports = new SpoHubSiteSetCommand();