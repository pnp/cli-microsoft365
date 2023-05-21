import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./tenant-applicationcustomizer-list');

describe(commands.TENANT_APPLICATIONCUSTOMIZER_LIST, () => {
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = 'https://contoso.sharepoint.com/sites/apps';
  const applicationCustomizerResponse = {
    value:
      [
        {
          "FileSystemObjectType": 0,
          "Id": 8,
          "ServerRedirectedEmbedUri": null,
          "ServerRedirectedEmbedUrl": "",
          "ID": 8,
          "ContentTypeId": "0x00693E2C487575B448BD420C12CEAE7EFE",
          "Title": "HelloWorld",
          "Modified": "2023-05-21T14:31:30Z",
          "Created": "2023-05-21T14:31:30Z",
          "AuthorId": 9,
          "EditorId": 9,
          "OData__UIVersionString": "1.0",
          "Attachments": false,
          "GUID": "23951a41-f613-440e-8119-8f1e87df1d1a",
          "OData__ColorTag": null,
          "ComplianceAssetId": null,
          "TenantWideExtensionComponentId": "d54e75e7-af4d-455f-9101-a5d906692ecd",
          "TenantWideExtensionComponentProperties": "{\"testMessage\":\"Test message\"}",
          "TenantWideExtensionWebTemplate": null,
          "TenantWideExtensionListTemplate": 0,
          "TenantWideExtensionLocation": "ClientSideExtension.ApplicationCustomizer",
          "TenantWideExtensionSequence": 0,
          "TenantWideExtensionHostProperties": null,
          "TenantWideExtensionDisabled": false
        }
      ]
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.spoUrl = spoUrl;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_APPLICATIONCUSTOMIZER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'TenantWideExtensionComponentId', 'TenantWideExtensionWebTemplate']);
  });

  it('throws error when tenant app catalog doesn\'t exist', async () => {
    const errorMessage = 'No app catalog URL found';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: null };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError(errorMessage));
  });

  it('retrieves application customizers that are installed tenant wide', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items`) {
        return applicationCustomizerResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(applicationCustomizerResponse.value));
  });

  it('correctly handles no tenant wide installed application customizers found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogToStderrSpy.calledWith('No tenant wide installed application customizers found'));
  });

  it('handles error when retrieving tenant wide installed application customizers', async () => {
    const errorMessage = 'An error has occurred';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError(errorMessage));
  });
});