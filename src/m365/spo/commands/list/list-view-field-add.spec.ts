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
import commands from '../../commands.js';
import command from './list-view-field-add.js';

describe(commands.LIST_VIEW_FIELD_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  const stubAllGetRequests: any = () => {
    return sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/fields/getbyinternalnameortitle') > -1 || (opts.url as string).indexOf('/fields/getbyid') > -1) {
        return {
          "AllowDisplay": true,
          "AllowMultipleValues": false,
          "AutoIndexed": false,
          "CanBeDeleted": false,
          "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
          "ClientSideComponentProperties": null,
          "CustomFormatter": null,
          "DefaultFormula": null,
          "DefaultValue": null,
          "DependentLookupInternalNames": [],
          "Description": "",
          "Direction": "none",
          "EnforceUniqueValues": false,
          "EntityPropertyName": "Author",
          "FieldTypeKind": 20,
          "Filterable": true,
          "FromBaseType": true,
          "Group": "Custom Columns",
          "Hidden": false,
          "Id": "1df5e554-ec7e-46a6-901d-d85a3881cb18",
          "Indexed": false,
          "InternalName": "Author",
          "IsDependentLookup": false,
          "IsRelationship": false,
          "JSLink": "clienttemplates.js",
          "LookupField": "",
          "LookupList": "{f978b511-305d-45e9-a7e7-f234a67e956d}",
          "LookupWebId": "c0950f14-23ce-4778-977a-9df11b866ede",
          "PinnedToFiltersPane": false,
          "Presence": true,
          "PrimaryFieldId": null,
          "ReadOnlyField": true,
          "RelationshipDeleteBehavior": 0,
          "Required": false,
          "SchemaXml": "<Field ID=\"{1df5e554-ec7e-46a6-901d-d85a3881cb18}\" ColName=\"tp_Author\" RowOrdinal=\"0\" ReadOnly=\"TRUE\" Type=\"User\" List=\"UserInfo\" Name=\"Author\" DisplayName=\"Created By\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Author\" FromBaseType\"TRUE\" />",
          "Scope": "/sites/ninja/Shared Documents",
          "Sealed": false,
          "SelectionGroup": 0,
          "SelectionMode": 1,
          "ShowInFiltersPane": 0,
          "Sortable": true,
          "StaticName": "Author",
          "Title": "Created By",
          "TypeAsString": "User",
          "TypeDisplayName": "Person or Group",
          "TypeShortDescription": "Person or Group",
          "UnlimitedLengthInDocumentLibrary": false,
          "ValidationFormula": null,
          "ValidationMessage": null
        };
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    requests = [];
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
    assert.strictEqual(command.name, commands.LIST_VIEW_FIELD_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the viewId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: '12345', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the position option is defined and is not a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', position: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('add the field by title to viewId and listTitle (debug)', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('add the field by title to viewId and listUrl', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('%2Fsites%2Fninja%2FShared%20Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl: '/sites/ninja/Shared Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/GetList('%2Fsites%2Fninja%2FShared%20Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('add the field by title from viewTitle and listTitle', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewTitle: 'MyView', title: 'Created By' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('add the field by title from viewTitle and listId', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', title: 'Created By' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('add the field by title to viewId and listId', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('add the field by id to viewId and listTitle', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('add the field by id from viewTitle and listTitle', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewTitle: 'MyView', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('add the field by id from viewTitle and listId', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('add the field by id to viewId and listId', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('move the field by title to the position index to viewId of listTitle (debug)', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === "https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/moveviewfieldto" ||
        opts.url === "https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/views('cc27a922-8224-4296-90a5-ebbc54da2e81')/viewfields/addviewfield('Author')") {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', title: 'Created By', position: 1 } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/viewfields/moveviewfieldto`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('move the field by title to the position index to viewTitle of listId', async () => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === "https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/moveviewfieldto" ||
        opts.url === "https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0cd891ef-afce-4e55-b836-fce03286cccf')/views/GetByTitle('MyView')/viewfields/addviewfield('Author')") {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', title: 'Created By', position: 1 } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/viewfields/moveviewfieldto`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('handles error correctly', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        viewId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        id: '330f29c5-5c4c-465f-9f4b-7903020ae1ce',
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        force: true
      }
    } as any), new CommandError(error.error['odata.error'].message.value));
  });
});
