import { Logger } from '../cli/Logger.js';
import request, { CliRequestOptions } from '../request.js';
import { spo } from './spo.js';

interface OrgAssetsLibrary {
  DisplayName: string;
  FileType: string;
  LibraryUrl: {
    DecodedUrl: string;
  };
  ListId: string;
  OrgAssetFlags: number;
  OrgAssetType: number;
  ThumbnailUrl: {
    DecodedUrl: string;
  };
  UniqueId: string;
}

interface OrgAssetsLibraryCollection {
  OrgAssetsLibraries: OrgAssetsLibrary[];
  Items: OrgAssetsLibrary[];
}

export interface BrandCenterConfiguration {
  BrandColorsListId: string;
  BrandColorsListUrl: string | null;
  BrandFontLibraryId: string;
  BrandFontLibraryUrl: string | null;
  IsBrandCenterSiteFeatureEnabled: boolean;
  IsPublicCdnEnabled: boolean;
  OrgAssets: {
    CentralAssetRepositoryLibraries?: OrgAssetsLibraryCollection;
    Domain: {
      DecodedUrl: string;
    };
    OrgAssetsLibraries: OrgAssetsLibraryCollection;
    SiteId: string;
    Url: {
      DecodedUrl: string;
    };
    WebId: string;
  };
  SiteId: string;
  SiteUrl: string;
}

export const brandCenter = {
  /**
   * Gets the brand center configuration for the specified site
   * @param logger Logger instance for verbose output
   * @param debug Debug flag for detailed logging
   * @returns Promise<BrandCenterConfiguration> Brand center configuration object
   */
  async getBrandCenterConfiguration(logger: Logger, debug: boolean = false): Promise<BrandCenterConfiguration> {
    if (debug) {
      await logger.logToStderr(`Retrieving brand center configuration...`);
    }

    const spoAdminUrl = await spo.getSpoAdminUrl(logger, debug);
    const brandConfigRequestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_api/SPO.Tenant/GetBrandCenterConfiguration`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const brandConfig = await request.get<BrandCenterConfiguration>(brandConfigRequestOptions);

    if (debug) {
      await logger.logToStderr(`Successfully retrieved brand center configuration`);
    }

    return brandConfig;
  }
};