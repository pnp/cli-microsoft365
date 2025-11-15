import { z } from 'zod';
import { Logger } from '../../../cli/Logger.js';
import { globalOptionsZod } from '../../../Command.js';
import config from "../../../config.js";
import request, { CliRequestOptions } from "../../../request.js";
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from "../../../utils/spo.js";
import SpoCommand from "../../base/SpoCommand.js";
import { SiteProperties } from "../../spo/commands/site/SiteProperties.js";
import { SPOSitePropertiesEnumerable } from "../../spo/commands/site/SPOSitePropertiesEnumerable.js";
import commands from "../commands.js";

export const options = z.strictObject({ ...globalOptionsZod.shape });

class OneDriveListCommand extends SpoCommand {
  private allSites?: SiteProperties[];

  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return "Retrieves a list of OneDrive sites";
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving list of OneDrive sites...`);
      }

      this.allSites = [];

      await this.getAllSites(spoAdminUrl, '0', undefined, logger);

      await logger.log(this.allSites);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getAllSites(spoAdminUrl: string, startIndex: string | undefined, formDigest: FormDigestInfo | undefined, logger: Logger): Promise<void> {
    const formDigestInfo: FormDigestInfo = await spo.ensureFormDigest(spoAdminUrl, logger, formDigest, this.debug);

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestInfo.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">1</Property><Property Name="StartIndex" Type="String">${startIndex}</Property><Property Name="Template" Type="String">SPSPERS</Property></Parameter></Parameters></Method></ObjectPaths></Request>`
    };

    const res: string = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      throw (response.ErrorInfo.ErrorMessage);
    }
    else {
      const sites: SPOSitePropertiesEnumerable = json[json.length - 1];
      this.allSites!.push(...sites._Child_Items_);

      if (sites.NextStartIndexFromSharePoint) {
        await this.getAllSites(spoAdminUrl, sites.NextStartIndexFromSharePoint, formDigest, logger);
      }
    }
  }
}

export default new OneDriveListCommand();