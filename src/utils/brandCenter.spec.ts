import assert from 'assert';
import sinon from 'sinon';
import { Logger } from '../cli/Logger.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { spo } from './spo.js';
import { brandCenter, BrandCenterConfiguration } from './brandCenter.js';

describe('utils/brandCenter', () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  const mockBrandCenterConfiguration: BrandCenterConfiguration = {
    BrandColorsListId: 'mock-colors-list-id',
    BrandColorsListUrl: 'https://contoso.sharepoint.com/sites/brandcenter/Lists/BrandColors',
    BrandFontLibraryId: 'mock-fonts-library-id',
    BrandFontLibraryUrl: 'https://contoso.sharepoint.com/sites/brandcenter/BrandFonts',
    IsBrandCenterSiteFeatureEnabled: true,
    IsPublicCdnEnabled: true,
    OrgAssets: {
      CentralAssetRepositoryLibraries: {
        OrgAssetsLibraries: [],
        Items: []
      },
      Domain: {
        DecodedUrl: 'https://contoso.sharepoint.com'
      },
      OrgAssetsLibraries: {
        OrgAssetsLibraries: [{
          DisplayName: 'Brand Assets',
          FileType: 'Image',
          LibraryUrl: {
            DecodedUrl: 'https://contoso.sharepoint.com/sites/brandcenter/BrandAssets'
          },
          ListId: 'mock-list-id',
          OrgAssetFlags: 1,
          OrgAssetType: 1,
          ThumbnailUrl: {
            DecodedUrl: 'https://contoso.sharepoint.com/sites/brandcenter/BrandAssets/thumbnail.jpg'
          },
          UniqueId: 'mock-unique-id'
        }],
        Items: []
      },
      SiteId: 'mock-site-id',
      Url: {
        DecodedUrl: 'https://contoso.sharepoint.com/sites/brandcenter'
      },
      WebId: 'mock-web-id'
    },
    SiteId: 'mock-site-id',
    SiteUrl: 'https://contoso.sharepoint.com/sites/brandcenter'
  };

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      spo.getSpoAdminUrl
    ]);
  });

  it('retrieves brand center configuration successfully', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(request, 'get').resolves(mockBrandCenterConfiguration);

    const result = await brandCenter.getBrandCenterConfiguration(logger, false);

    assert.deepStrictEqual(result, mockBrandCenterConfiguration);
  });

  it('retrieves brand center configuration with debug logging', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(request, 'get').resolves(mockBrandCenterConfiguration);

    await brandCenter.getBrandCenterConfiguration(logger, true);

    assert(loggerLogToStderrSpy.calledWith('Retrieving brand center configuration...'));
    assert(loggerLogToStderrSpy.calledWith('Successfully retrieved brand center configuration'));
  });

  it('handles errors without debug logging when debug is false', async () => {
    const errorMessage = 'Access denied';
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(
      brandCenter.getBrandCenterConfiguration(logger, false),
      { message: errorMessage }
    );

    assert(loggerLogToStderrSpy.notCalled);
  });

  it('returns configuration with IsBrandCenterSiteFeatureEnabled false when feature is disabled', async () => {
    const disabledConfig: BrandCenterConfiguration = {
      ...mockBrandCenterConfiguration,
      IsBrandCenterSiteFeatureEnabled: false
    };

    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(request, 'get').resolves(disabledConfig);

    const result = await brandCenter.getBrandCenterConfiguration(logger, false);

    assert.deepStrictEqual(result, disabledConfig);
  });

  it('handles response with null values', async () => {
    const configWithNulls: BrandCenterConfiguration = {
      ...mockBrandCenterConfiguration,
      BrandColorsListUrl: null,
      BrandFontLibraryUrl: null
    };

    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(request, 'get').resolves(configWithNulls);

    const result = await brandCenter.getBrandCenterConfiguration(logger, false);

    assert.deepStrictEqual(result, configWithNulls);
  });
});