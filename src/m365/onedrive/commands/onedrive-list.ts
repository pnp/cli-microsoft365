import { Logger } from "../../../cli";
import config from "../../../config";
import request from "../../../request";
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from "../../../utils";
import SpoCommand from "../../base/SpoCommand";
import { SiteProperties } from "../../spo/commands/site/SiteProperties";
import { SPOSitePropertiesEnumerable } from "../../spo/commands/site/SPOSitePropertiesEnumerable";
import commands from "../commands";

class OneDriveListCommand extends SpoCommand {
  private allSites?: SiteProperties[];
  
  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return "Retrieves a list of OneDrive sites";
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url'];
  }

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {
    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((spoAdminUrl: string): Promise<void> => {
        if (this.verbose) {
          logger.logToStderr(`Retrieving list of OneDrive sites...`);
        }

        this.allSites = [];

        return this.getAllSites(spoAdminUrl, '0', undefined, logger);
      })
      .then(_ => {
        logger.log(this.allSites);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  private getAllSites(spoAdminUrl: string, startIndex: string | undefined, formDigest: FormDigestInfo | undefined, logger: Logger): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      spo
        .ensureFormDigest(spoAdminUrl, logger, formDigest, this.debug)
        .then((res: FormDigestInfo): Promise<string> => {
          const requestOptions: any = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': res.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">1</Property><Property Name="StartIndex" Type="String">${startIndex}</Property><Property Name="Template" Type="String">SPSPERS</Property></Parameter></Parameters></Method></ObjectPaths></Request>`
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            reject(response.ErrorInfo.ErrorMessage);
            return;
          }
          else {
            const sites: SPOSitePropertiesEnumerable = json[json.length - 1];
            this.allSites!.push(...sites._Child_Items_);

            if (sites.NextStartIndexFromSharePoint) {
              this
                .getAllSites(spoAdminUrl, sites.NextStartIndexFromSharePoint, formDigest, logger)
                .then(_ => resolve(), err => reject(err));
            }
            else {
              resolve();
            }
          }
        }, err => reject(err));
    });
  }
}

module.exports = new OneDriveListCommand();