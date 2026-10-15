import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { HubSiteProperties } from './HubSiteProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID`
  }).alias('i'),
  title: z.string().optional().alias('t'),
  description: z.string().optional().alias('d'),
  logoUrl: z.string().optional().alias('l')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoHubSiteSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_SET;
  }

  public get description(): string {
    return 'Updates properties of the specified hub site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.title !== undefined || opts.description !== undefined || opts.logoUrl !== undefined, {
        error: 'Specify title, description or logoUrl to update'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      if (this.verbose) {
        await logger.logToStderr(`Updating hub site ${args.options.id}...`);
      }

      const title: string = typeof args.options.title === 'string' ? `<SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">${formatting.escapeXml(args.options.title)}</Parameter></SetProperty>` : '';
      const description: string = typeof args.options.description === 'string' ? `<SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">${formatting.escapeXml(args.options.description)}</Parameter></SetProperty>` : '';
      const logoUrl: string = typeof args.options.logoUrl === 'string' ? `<SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">${formatting.escapeXml(args.options.logoUrl)}</Parameter></SetProperty>` : '';

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query>${title}${logoUrl}${description}<Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const hubSite: HubSiteProperties = json.pop();
        delete hubSite._ObjectType_;

        hubSite.ID = hubSite.ID.replace('/Guid(', '').replace(')/', '');
        hubSite.SiteId = hubSite.SiteId.replace('/Guid(', '').replace(')/', '');

        await logger.log(hubSite);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoHubSiteSetCommand();