import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-commandset-set.js';

describe(commands.TENANT_COMMANDSET_SET, () => {
  const id = 1;
  const clientSideComponentId = '9748c81b-d72e-4048-886a-e98649543743';
  const clientSideComponentProperties = '{ "someProperty": "Some value" }';
  const title = 'Some Command Set';
  const webTemplate = 'GROUP#0';
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = `https://contoso.sharepoint.com/sites/apps`;
  const commandSetResponse = {
    "FileSystemObjectType": 0,
    "Id": id,
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
    "GUID": "6e6f2429-cdec-4b90-89da-139d2665919e",
    "ComplianceAssetId": null,
    "TenantWideExtensionComponentId": clientSideComponentId,
    "TenantWideExtensionComponentProperties": "{\"testMessage\":\"Test message\"}",
    "TenantWideExtensionWebTemplate": null,
    "TenantWideExtensionListTemplate": 101,
    "TenantWideExtensionLocation": "ClientSideExtension.ListViewCommandSet.ContextMenu",
    "TenantWideExtensionSequence": 0,
    "TenantWideExtensionHostProperties": null,
    "TenantWideExtensionDisabled": false
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = spoUrl;
    commandInfo = cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_COMMANDSET_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates a tenant-wide ListView Command Set for lists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "Title",
              FieldValue: title,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: clientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: clientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionListTemplate",
              FieldValue: 100,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet",
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionWebTemplate",
              FieldValue: webTemplate,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentProperties",
              FieldValue: clientSideComponentProperties,
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, { options: { id: id, newTitle: title, clientSideComponentId: clientSideComponentId, listType: 'List', location: 'Both', webTemplate: webTemplate, clientSideComponentProperties: clientSideComponentProperties, verbose: true } }));
  });

  it('updates a tenant-wide ListView Command Set for lists with location ContextMenu and listType SitePages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionListTemplate",
              FieldValue: 119,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet.ContextMenu",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, { options: { id: id, location: 'ContextMenu', listType: 'SitePages' } }));
  });

  it('updates a tenant-wide ListView Command Set for lists with location CommandBar and listType Library ', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionListTemplate",
              FieldValue: 101,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet.CommandBar",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, { options: { id: id, location: 'CommandBar', listType: 'Library' } }));
  });

  const errorMessage = 'No app catalog URL found';

  it('throws error when tenant app catalog doesn\'t exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: null };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        newTitle: title
      }
    }), new CommandError(errorMessage));
  });

  it('throws error when retrieving a tenant app catalog fails with an exception', async () => {
    const errorMessage = 'Couldn\'t retrieve tenant app catalog URL';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        newTitle: title
      }
    }), new CommandError(errorMessage));
  });

  it('throws error when retrieving an item which is not a listview commandset', async () => {
    const errorMessage = 'The item is not a ListViewCommandSet';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return {
          "FileSystemObjectType": 0,
          "ID": 4,
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
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        newTitle: title
      }
    }), new CommandError(errorMessage));
  });

  it('fails validation if no option to update is specified is not a valid Guid', async () => {
    const actual = await command.validate({ options: { id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if clientSideComponentId is not a valid Guid', async () => {
    const actual = await command.validate({ options: { id: id, clientSideComponentId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if location value is not valid', async () => {
    const actual = await command.validate({ options: { id: id, location: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listType value is not valid', async () => {
    const actual = await command.validate({ options: { id: id, listType: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all properties are specified', async () => {
    const actual = await command.validate({ options: { id: id, newTitle: title, clientSideComponentId: clientSideComponentId, listType: 'List', location: 'Both', webTemplate: webTemplate, clientSideComponentProperties: clientSideComponentProperties } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});