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
import { GetExternalUsersResults } from './GetExternalUsersResults.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  filter: z.string().optional(),
  pageSize: z.string().refine(val => {
    const num = parseInt(val);
    return !isNaN(num) && num >= 1 && num <= 50;
  }, {
    error: e => `pageSize must be between 1 and 50. The specified value ${e.input} is not valid.`
  }).optional().alias('p'),
  position: z.string().refine(val => {
    const num = parseInt(val);
    return !isNaN(num) && num >= 0;
  }, {
    error: e => `position must be greater than or equal to 0. The specified value ${e.input} is not valid.`
  }).optional().alias('i'),
  sortOrder: z.enum(['asc', 'desc']).optional().alias('s'),
  siteUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).optional().alias('u')
});

export type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoExternalUserListCommand extends SpoCommand {
  public get name(): string {
    return commands.EXTERNALUSER_LIST;
  }

  public get description(): string {
    return 'Lists external users in the tenant';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving information about external users...`);
      }

      const position: number = parseInt(args.options.position || '0');
      const pageSize: number = parseInt(args.options.pageSize || '10');
      const sortOrder: number = args.options.sortOrder === 'desc' ? 1 : 0;

      let payload: string = '';

      if (args.options.siteUrl) {
        payload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="135" ObjectPathId="134" /><Query Id="136" ObjectPathId="134"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="134" ParentId="131" Name="GetExternalUsersForSite"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.siteUrl)}</Parameter><Parameter Type="Int32">${position}</Parameter><Parameter Type="Int32">${pageSize}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.filter || '')}</Parameter><Parameter Type="Enum">${sortOrder}</Parameter></Parameters></Method><Constructor Id="131" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`;
      }
      else {
        payload = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="109" ObjectPathId="108" /><Query Id="110" ObjectPathId="108"><Query SelectAllProperties="false"><Properties><Property Name="TotalUserCount" ScalarProperty="true" /><Property Name="UserCollectionPosition" ScalarProperty="true" /><Property Name="ExternalUserCollection"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="DisplayName" ScalarProperty="true" /><Property Name="InvitedAs" ScalarProperty="true" /><Property Name="UniqueId" ScalarProperty="true" /><Property Name="AcceptedAs" ScalarProperty="true" /><Property Name="WhenCreated" ScalarProperty="true" /><Property Name="InvitedBy" ScalarProperty="true" /></Properties></ChildItemQuery></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="108" ParentId="105" Name="GetExternalUsers"><Parameters><Parameter Type="Int32">${position}</Parameter><Parameter Type="Int32">${pageSize}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.filter || '')}</Parameter><Parameter Type="Enum">${sortOrder}</Parameter></Parameters></Method><Constructor Id="105" TypeId="{e45fd516-a408-4ca4-b6dc-268e2f1f0f83}" /></ObjectPaths></Request>`;
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: payload
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const results: GetExternalUsersResults = json.pop();

        if (results.TotalUserCount > 0) {
          await logger.log(results.ExternalUserCollection._Child_Items_.map(e => {
            delete e._ObjectType_;
            const dateChunks: number[] = (e.WhenCreated as string)
              .replace('/Date(', '')
              .replace(')/', '')
              .split(',')
              .map(c => {
                return parseInt(c);
              });
            e.WhenCreated = new Date(dateChunks[0], dateChunks[1], dateChunks[2], dateChunks[3], dateChunks[4], dateChunks[5], dateChunks[6]);
            return e;
          }));
        }
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoExternalUserListCommand();