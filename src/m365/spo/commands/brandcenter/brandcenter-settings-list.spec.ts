import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { z } from 'zod';
import commands from '../../commands.js';
import command from './brandcenter-settings-list.js';

describe(commands.BRANDCENTER_SETTINGS_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  const successResponse = {
    "BrandColorsListId": "00000000-0000-0000-0000-000000000000",
    "BrandColorsListUrl": null,
    "BrandFontLibraryId": "23af51de-856c-4d00-aa11-0d03af0e46e3",
    "BrandFontLibraryUrl": {
      "DecodedUrl": "https://contoso.sharepoint.com/sites/BrandGuide/Fonts"
    },
    "IsBrandCenterSiteFeatureEnabled": true,
    "IsPublicCdnEnabled": true,
    "OrgAssets": {
      "CentralAssetRepositoryLibraries": null,
      "Domain": {
        "DecodedUrl": "https://contoso.sharepoint.com"
      },
      "OrgAssetsLibraries": {
        "OrgAssetsLibraries": [
          {
            "DisplayName": "Fonts",
            "FileType": "",
            "LibraryUrl": {
              "DecodedUrl": "sites/BrandGuide/Fonts"
            },
            "ListId": "23af51de-856c-4d00-aa11-0d03af0e46e3",
            "OrgAssetFlags": 0,
            "OrgAssetType": 8,
            "ThumbnailUrl": null,
            "UniqueId": "00000000-0000-0000-0000-000000000000"
          }
        ],
        "Items": [
          {
            "DisplayName": "Fonts",
            "FileType": "",
            "LibraryUrl": {
              "DecodedUrl": "sites/BrandGuide/Fonts"
            },
            "ListId": "23af51de-856c-4d00-aa11-0d03af0e46e3",
            "OrgAssetFlags": 0,
            "OrgAssetType": 8,
            "ThumbnailUrl": null,
            "UniqueId": "00000000-0000-0000-0000-000000000000"
          }
        ]
      },
      "SiteId": "52b46e48-9c0c-40cb-a955-13eb6c717ff3",
      "Url": {
        "DecodedUrl": "/sites/BrandGuide"
      },
      "WebId": "206988d5-e133-4a24-819d-24101f3407ce"
    },
    "SiteId": "52b46e48-9c0c-40cb-a955-13eb6c717ff3",
    "SiteUrl": "https://contoso.sharepoint.com/sites/BrandGuide"
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
  });

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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.BRANDCENTER_SETTINGS_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ option: "value" });
    assert.strictEqual(actual.success, false);
  });

  it('makes correct GET request to retrieve brand center settings', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Brandcenter/Configuration`) {
        return successResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert.strictEqual(getStub.firstCall.args[0].url, 'https://contoso.sharepoint.com/_api/Brandcenter/Configuration');
  });

  it('successfully logs brand center settings', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Brandcenter/Configuration`) {
        return successResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(successResponse));
  });

  it('correctly handles error when retrieving settings', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "accessDenied",
        "message": "Access denied"
      }
    });

    await assert.rejects(command.action(logger, { options: {} }),
      new CommandError('Access denied'));
  });
});