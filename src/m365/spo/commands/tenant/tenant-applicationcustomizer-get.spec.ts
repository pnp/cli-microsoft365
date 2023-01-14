import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as SpoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get';
import * as SpoListItemListCommand from '../listitem/listitem-list';
const command: Command = require('./tenant-applicationcustomizer-get');

describe(commands.TENANT_APPLICATIONCUSTOMIZER_GET, () => {
  const title = 'Some customizer';
  const id = '14125658-a9bc-4ddf-9c75-1b5767c9a337';
  const clientSideComponentId = '7096cded-b83d-4eab-96f0-df477ed7c0bc';
  const appCatalogUrl = 'https://contoso.sharepoint.com/sites/apps';
  const applicationCustomizerResponse: any = [{
    "FileSystemObjectType": 0,
    "Id": 4,
    "ServerRedirectedEmbedUri": null,
    "ServerRedirectedEmbedUrl": "",
    "ContentTypeId": "0x00693E2C487575B448BD420C12CEAE7EFE",
    "Title": title,
    "Modified": "2023-01-11T15:47:38Z",
    "Created": "2023-01-11T15:47:38Z",
    "AuthorId": 9,
    "EditorId": 9,
    "OData__UIVersionString": "1.0",
    "Attachments": false,
    "GUID": id,
    "ComplianceAssetId": null,
    "TenantWideExtensionComponentId": clientSideComponentId,
    "TenantWideExtensionComponentProperties": "{\"testMessage\":\"Test message\"}",
    "TenantWideExtensionWebTemplate": null,
    "TenantWideExtensionListTemplate": 0,
    "TenantWideExtensionLocation": "ClientSideExtension.ApplicationCustomizer",
    "TenantWideExtensionSequence": 0,
    "TenantWideExtensionHostProperties": null,
    "TenantWideExtensionDisabled": false
  }];

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_APPLICATIONCUSTOMIZER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientSideComponentId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { clientSideComponentId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('handles error when app catalog not registered', async () => {
    const errorMessage = 'No app catalog URL found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoTenantAppCatalogUrlGetCommand) {
        return '';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        title: title
      }
    }), new CommandError(errorMessage));
  });

  it('retrieves an application customizer by title', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoTenantAppCatalogUrlGetCommand) {
        return { stdout: appCatalogUrl };
      }

      if (command === SpoListItemListCommand) {
        return { stdout: JSON.stringify(applicationCustomizerResponse) };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        title: title
      }
    });
    assert(loggerLogSpy.calledWith(applicationCustomizerResponse[0]));
  });

  it('retrieves an application customizer by id', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoTenantAppCatalogUrlGetCommand) {
        return { stdout: appCatalogUrl };
      }

      if (command === SpoListItemListCommand) {
        return { stdout: JSON.stringify(applicationCustomizerResponse) };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: id
      }
    });
    assert(loggerLogSpy.calledWith(applicationCustomizerResponse[0]));
  });

  it('retrieves an application customizer by clientSideComponentId', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoTenantAppCatalogUrlGetCommand) {
        return { stdout: appCatalogUrl };
      }

      if (command === SpoListItemListCommand) {
        return { stdout: JSON.stringify(applicationCustomizerResponse) };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId
      }
    });
    assert(loggerLogSpy.calledWith(applicationCustomizerResponse[0]));
  });

  it('handles error when specified application customizer not found', async () => {
    const errorMessage = 'The specified application customizer was not found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoTenantAppCatalogUrlGetCommand) {
        return { stdout: appCatalogUrl };
      }

      if (command === SpoListItemListCommand) {
        return { stdout: JSON.stringify([]) };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title
      }
    }), new CommandError(errorMessage));
  });
});