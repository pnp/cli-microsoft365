import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-applicationcustomizer-list.js';
import { ListItemListOptions, spoListItem } from '../../../../utils/spoListItem.js';
import { spo } from '../../../../utils/spo.js';

describe(commands.TENANT_APPLICATIONCUSTOMIZER_LIST, () => {
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = 'https://contoso.sharepoint.com/sites/apps';
  const applicationCustomizer = {
    "FileSystemObjectType": 0,
    "Id": 8,
    "ServerRedirectedEmbedUri": null,
    "ServerRedirectedEmbedUrl": "",
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
  };

  const applicationCustomizerList = [
    { ...applicationCustomizer }
  ];

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = spoUrl;
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
      spo.getTenantAppCatalogUrl,
      spoListItem.getListItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
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

    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(null);

    await assert.rejects(command.action(logger, { options: {} }), new CommandError(errorMessage));
  });

  it('retrieves application customizers that are installed tenant wide', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer'`
        ) {
          return applicationCustomizerList as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([applicationCustomizer]));
  });

  it('handles error when retrieving tenant wide installed application customizers', async () => {
    const errorMessage = 'An error has occurred';

    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer'`
        ) {
          throw errorMessage;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError(errorMessage));
  });
});