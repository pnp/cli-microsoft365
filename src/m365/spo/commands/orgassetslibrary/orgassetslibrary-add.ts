import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  libraryUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, {
    message: 'The value of the libraryUrl is not a valid SharePoint Online site URL.'
  }),
  thumbnailUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, {
    message: 'The value of the thumbnailUrl is not a valid SharePoint Online site URL.'
  }).optional(),
  cdnType: z.enum(['Public', 'Private']).optional(),
  orgAssetType: z.enum(['ImageDocumentLibrary', 'OfficeTemplateLibrary', 'OfficeFontLibrary']).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

enum OrgAssetType {
  ImageDocumentLibrary = 1,
  OfficeTemplateLibrary = 2,
  OfficeFontLibrary = 4
}

class SpoOrgAssetsLibraryAddCommand extends SpoCommand {
  public get name(): string {
    return commands.ORGASSETSLIBRARY_ADD;
  }

  public get description(): string {
    return 'Promotes an existing library to become an organization assets library';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const cdnTypeString: string = args.options.cdnType || 'Private';
    const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;
    const thumbnailSchema: string = typeof args.options.thumbnailUrl === 'undefined' ? `<Parameter Type="Null" />` : `<Parameter Type="String">${args.options.thumbnailUrl}</Parameter>`;

    try {
      const orgAssetType = this.getOrgAssetType(args.options.orgAssetType);
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="String">${args.options.libraryUrl}</Parameter>${thumbnailSchema}<Parameter Type="Enum">${orgAssetType}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private getOrgAssetType(orgAssetType: string | undefined): OrgAssetType {
    switch (orgAssetType) {
      case 'OfficeTemplateLibrary':
        return OrgAssetType.OfficeTemplateLibrary;
      case 'OfficeFontLibrary':
        return OrgAssetType.OfficeFontLibrary;
      default:
        return OrgAssetType.ImageDocumentLibrary;
    }
  }
}

export default new SpoOrgAssetsLibraryAddCommand();
