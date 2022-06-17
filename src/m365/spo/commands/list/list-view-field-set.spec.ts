import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./list-view-field-set');

describe(commands.LIST_VIEW_FIELD_SET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  const stubAllGetRequests: any = () => {
    return sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/fields/getbyinternalnameortitle') > -1 || (opts.url as string).indexOf('/fields/getbyid') > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject('Invalid request');
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_VIEW_FIELD_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('move the field by title to the position index to viewTitle of listId', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/viewfields/moveviewfieldto`) > -1 &&
        JSON.stringify(opts.data) === `{"field":"Author","index":1}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', fieldTitle: 'Created By', fieldPosition: 1 } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('move the field by title to the position index to viewId of listTitle (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/viewfields/moveviewfieldto`) > -1 &&
        JSON.stringify(opts.data) === `{"field":"Author","index":1}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldTitle: 'Created By', fieldPosition: 1 } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('move the field by id to the position index to viewId of listTitle', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/viewfields/moveviewfieldto`) > -1 &&
        JSON.stringify(opts.data) === `{"field":"Author","index":1}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '1df5e554-ec7e-46a6-901d-d85a3881cb18', fieldPosition: 1 } }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('move the field by id to the position index to viewTitle of listId (debug)', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/viewfields/moveviewfieldto`) > -1 &&
        JSON.stringify(opts.data) === `{"field":"Author","index":1}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'MyView', fieldId: '1df5e554-ec7e-46a6-901d-d85a3881cb18', fieldPosition: 1 } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when list id option is passed', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        viewId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce',
        webUrl: 'https://contoso.sharepoint.com',
        listId: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        fieldPosition: 1
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when list title option is passed', (done) => {
    stubAllGetRequests();

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        viewId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce',
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        fieldPosition: 1
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if one of listId or listTitle options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if one of viewId or viewTitle options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if one of fieldId or fieldTitle options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fieldPosition option is not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e85' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e85', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the viewId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: '12345', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fieldId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '12345', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the viewId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the fieldId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', listTitle: 'Documents', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both viewId and viewTitle options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'My view', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both fieldId and fieldTitle options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewTitle: 'My view', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldTitle: 'Created By', fieldPosition: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fieldPosition option is defined and is not a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the fieldPosition option is defined and is a valid number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0cd891ef-afce-4e55-b836-fce03286cccf', viewId: 'cc27a922-8224-4296-90a5-ebbc54da2e81', fieldId: '330f29c5-5c4c-465f-9f4b-7903020ae1ce', fieldPosition: 1 } }, commandInfo);
    assert(actual);
  });
});