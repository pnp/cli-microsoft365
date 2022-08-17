import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./field-list');

describe(commands.FIELD_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.FIELD_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'Title', 'Group', 'Hidden']);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'site.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the list ID is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list id', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list title', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list url', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if title and id are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and id and url are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and url are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
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
        assert(loggerLogSpy.calledWith([
          {
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
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all list columns from list queried by title', (done) => {
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
        assert(loggerLogSpy.calledWith([
          {
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
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all list columns from list queried by url', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2Ftest')/fields`) > -1) {
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

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'test' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
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
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all list columns from list queried by id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'3c0e9e00-8fcc-479f-9d8d-3447cda34c5b')/fields`) > -1) {
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

    command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: '3c0e9e00-8fcc-479f-9d8d-3447cda34c5b' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
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
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});