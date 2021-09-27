import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./field-list');

describe(commands.FIELD_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FIELD_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves all site columns', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields`) > -1) {
        return Promise.resolve({
          "value": [{
            "AutoIndexed": false,
            "CanBeDeleted": true,
            "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
            "ClientSideComponentProperties": null,
            "ClientValidationFormula": null,
            "ClientValidationMessage": null,
            "CustomFormatter": null,
            "DefaultFormula": null,
            "DefaultValue": null,
            "Description": "",
            "Direction": "none",
            "EnforceUniqueValues": false,
            "EntityPropertyName": "fieldname",
            "FieldTypeKind": 2,
            "Filterable": true,
            "FromBaseType": false,
            "Group": "Core Contact and Calendar Columns",
            "Hidden": false,
            "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
            "IndexStatus": 0,
            "Indexed": false,
            "InternalName": "fieldname",
            "IsModern": false,
            "JSLink": "clienttemplates.js",
            "MaxLength": 255,
            "PinnedToFiltersPane": false,
            "ReadOnlyField": false,
            "Required": false,
            "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
            "Scope": "/sites/portal",
            "Sealed": true,
            "ShowInFiltersPane": 0,
            "Sortable": true,
            "StaticName": "fieldname",
            "Title": "Field Name",
            "TypeAsString": "Text",
            "TypeDisplayName": "Single line of text",
            "TypeShortDescription": "Single line of text",
            "ValidationFormula": null,
            "ValidationMessage": null
          }]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "value": [{
            "AutoIndexed": false,
            "CanBeDeleted": true,
            "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
            "ClientSideComponentProperties": null,
            "ClientValidationFormula": null,
            "ClientValidationMessage": null,
            "CustomFormatter": null,
            "DefaultFormula": null,
            "DefaultValue": null,
            "Description": "",
            "Direction": "none",
            "EnforceUniqueValues": false,
            "EntityPropertyName": "fieldname",
            "FieldTypeKind": 2,
            "Filterable": true,
            "FromBaseType": false,
            "Group": "Core Contact and Calendar Columns",
            "Hidden": false,
            "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
            "IndexStatus": 0,
            "Indexed": false,
            "InternalName": "fieldname",
            "IsModern": false,
            "JSLink": "clienttemplates.js",
            "MaxLength": 255,
            "PinnedToFiltersPane": false,
            "ReadOnlyField": false,
            "Required": false,
            "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
            "Scope": "/sites/portal",
            "Sealed": true,
            "ShowInFiltersPane": 0,
            "Sortable": true,
            "StaticName": "fieldname",
            "Title": "Field Name",
            "TypeAsString": "Text",
            "TypeDisplayName": "Single line of text",
            "TypeShortDescription": "Single line of text",
            "ValidationFormula": null,
            "ValidationMessage": null
          }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all list columns from list', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/fields`) > -1) {
        return Promise.resolve(
          {
            "value": [{
              "AutoIndexed": false,
              "CanBeDeleted": true,
              "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
              "ClientSideComponentProperties": null,
              "ClientValidationFormula": null,
              "ClientValidationMessage": null,
              "CustomFormatter": null,
              "DefaultFormula": null,
              "DefaultValue": null,
              "Description": "",
              "Direction": "none",
              "EnforceUniqueValues": false,
              "EntityPropertyName": "fieldname",
              "FieldTypeKind": 2,
              "Filterable": true,
              "FromBaseType": false,
              "Group": "Core Contact and Calendar Columns",
              "Hidden": false,
              "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
              "IndexStatus": 0,
              "Indexed": false,
              "InternalName": "fieldname",
              "IsModern": false,
              "JSLink": "clienttemplates.js",
              "MaxLength": 255,
              "PinnedToFiltersPane": false,
              "ReadOnlyField": false,
              "Required": false,
              "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
              "Scope": "/sites/portal/Documents",
              "Sealed": true,
              "ShowInFiltersPane": 0,
              "Sortable": true,
              "StaticName": "fieldname",
              "Title": "Field Name",
              "TypeAsString": "Text",
              "TypeDisplayName": "Single line of text",
              "TypeShortDescription": "Single line of text",
              "ValidationFormula": null,
              "ValidationMessage": null
            }]
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "value": [{
            "AutoIndexed": false,
            "CanBeDeleted": true,
            "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
            "ClientSideComponentProperties": null,
            "ClientValidationFormula": null,
            "ClientValidationMessage": null,
            "CustomFormatter": null,
            "DefaultFormula": null,
            "DefaultValue": null,
            "Description": "",
            "Direction": "none",
            "EnforceUniqueValues": false,
            "EntityPropertyName": "fieldname",
            "FieldTypeKind": 2,
            "Filterable": true,
            "FromBaseType": false,
            "Group": "Core Contact and Calendar Columns",
            "Hidden": false,
            "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
            "IndexStatus": 0,
            "Indexed": false,
            "InternalName": "fieldname",
            "IsModern": false,
            "JSLink": "clienttemplates.js",
            "MaxLength": 255,
            "PinnedToFiltersPane": false,
            "ReadOnlyField": false,
            "Required": false,
            "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
            "Scope": "/sites/portal/Documents",
            "Sealed": true,
            "ShowInFiltersPane": 0,
            "Sortable": true,
            "StaticName": "fieldname",
            "Title": "Field Name",
            "TypeAsString": "Text",
            "TypeDisplayName": "Single line of text",
            "TypeShortDescription": "Single line of text",
            "ValidationFormula": null,
            "ValidationMessage": null
          }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call the correct GET url when  list url specified', (done) => {
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal',  listUrl: 'Lists/Events' } }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/GetList(\'%2Fsites%2Fportal%2FLists%2FEvents\')/fields');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call the correct GET url when list title specified (verbose)', (done) => {
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, verbose: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle(\'Documents\')/fields');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call the correct GET url when field title and list title specified', (done) => {
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists/getByTitle(\'Documents\')/fields');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call the correct GET url when field title and list url specified', (done) => {
    const getStub = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists`) > -1) {
        return Promise.resolve({
          "Id": "03e45e84-1992-4d42-9116-26f756012634"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal',  listId: '03e45e84-1992-4d42-9116-26f756012634' } }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/portal/_api/web/lists(guid\'03e45e84-1992-4d42-9116-26f756012634\')/fields');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  

  it('correctly handles list not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/fields`) > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-1, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = command.validate({ options: { webUrl: 'site.com' } });
    assert.notStrictEqual(actual, true);
  });
 

  it('fails validation if the list ID is not a valid GUID', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales'} });
    assert.strictEqual(actual, true);
  });
});