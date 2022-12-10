import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import commands from '../../commands';
const command: Command = require('./field-add');

describe(commands.FIELD_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
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
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FIELD_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates site column using XML with the default options', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/CreateFieldAsXml`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 0
          }
        })) {
        return Promise.resolve({
          "AutoIndexed": false,
          "CanBeDeleted": true,
          "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
          "ClientSideComponentProperties": null,
          "CustomFormatter": null,
          "DefaultFormula": null,
          "DefaultValue": "[today]",
          "Description": "",
          "Direction": "none",
          "EnforceUniqueValues": false,
          "EntityPropertyName": "PnPAlertStartDateTime",
          "Filterable": true,
          "FromBaseType": false,
          "Group": "PnP Columns",
          "Hidden": false,
          "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
          "Indexed": false,
          "InternalName": "PnPAlertStartDateTime",
          "JSLink": "clienttemplates.js",
          "PinnedToFiltersPane": false,
          "ReadOnlyField": false,
          "Required": false,
          "SchemaXml": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"1\"><Default>[today]</Default></Field>",
          "Scope": "/sites/portal",
          "Sealed": false,
          "ShowInFiltersPane": 0,
          "Sortable": true,
          "StaticName": "PnPAlertStartDateTime",
          "Title": "Start date-time",
          "FieldTypeKind": 4,
          "TypeAsString": "DateTime",
          "TypeDisplayName": "Date and Time",
          "TypeShortDescription": "Date and Time",
          "ValidationFormula": null,
          "ValidationMessage": null,
          "DateTimeCalendarType": 0,
          "DisplayFormat": 1,
          "FriendlyDisplayFormat": 1
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>' } });
    assert(loggerLogSpy.calledWith({
      "AutoIndexed": false,
      "CanBeDeleted": true,
      "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
      "ClientSideComponentProperties": null,
      "CustomFormatter": null,
      "DefaultFormula": null,
      "DefaultValue": "[today]",
      "Description": "",
      "Direction": "none",
      "EnforceUniqueValues": false,
      "EntityPropertyName": "PnPAlertStartDateTime",
      "Filterable": true,
      "FromBaseType": false,
      "Group": "PnP Columns",
      "Hidden": false,
      "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
      "Indexed": false,
      "InternalName": "PnPAlertStartDateTime",
      "JSLink": "clienttemplates.js",
      "PinnedToFiltersPane": false,
      "ReadOnlyField": false,
      "Required": false,
      "SchemaXml": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"1\"><Default>[today]</Default></Field>",
      "Scope": "/sites/portal",
      "Sealed": false,
      "ShowInFiltersPane": 0,
      "Sortable": true,
      "StaticName": "PnPAlertStartDateTime",
      "Title": "Start date-time",
      "FieldTypeKind": 4,
      "TypeAsString": "DateTime",
      "TypeDisplayName": "Date and Time",
      "TypeShortDescription": "Date and Time",
      "ValidationFormula": null,
      "ValidationMessage": null,
      "DateTimeCalendarType": 0,
      "DisplayFormat": 1,
      "FriendlyDisplayFormat": 1
    }));
  });

  it('creates list column using XML with the DefaultValue option (debug)', async () => {
    const output = {
      "AutoIndexed": false,
      "CanBeDeleted": true,
      "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
      "ClientSideComponentProperties": null,
      "CustomFormatter": null,
      "DefaultFormula": null,
      "DefaultValue": "[today]",
      "Description": "",
      "Direction": "none",
      "EnforceUniqueValues": false,
      "EntityPropertyName": "PnPAlertStartDateTime",
      "Filterable": true,
      "FromBaseType": false,
      "Group": "PnP Columns",
      "Hidden": false,
      "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
      "Indexed": false,
      "InternalName": "PnPAlertStartDateTime",
      "JSLink": "clienttemplates.js",
      "PinnedToFiltersPane": false,
      "ReadOnlyField": false,
      "Required": false,
      "SchemaXml": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"1\"><Default>[today]</Default></Field>",
      "Scope": "/sites/portal",
      "Sealed": false,
      "ShowInFiltersPane": 0,
      "Sortable": true,
      "StaticName": "PnPAlertStartDateTime",
      "Title": "Start date-time",
      "FieldTypeKind": 4,
      "TypeAsString": "DateTime",
      "TypeDisplayName": "Date and Time",
      "TypeShortDescription": "Date and Time",
      "ValidationFormula": null,
      "ValidationMessage": null,
      "DateTimeCalendarType": 0,
      "DisplayFormat": 1,
      "FriendlyDisplayFormat": 1
    };
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Events')/fields/CreateFieldAsXml`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 0
          }
        })) {
        return Promise.resolve(output);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Events', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'DefaultValue' } });
    assert(loggerLogSpy.calledWith(output));
  });

  it('creates site column using XML with the AddToAllContentTypes, AddFieldToDefaultView, AddFieldCheckDisplayName options', async () => {
    const response = {
      "AutoIndexed": false,
      "CanBeDeleted": true,
      "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
      "ClientSideComponentProperties": null,
      "CustomFormatter": null,
      "DefaultFormula": null,
      "DefaultValue": "[today]",
      "Description": "",
      "Direction": "none",
      "EnforceUniqueValues": false,
      "EntityPropertyName": "PnPAlertStartDateTime",
      "Filterable": true,
      "FromBaseType": false,
      "Group": "PnP Columns",
      "Hidden": false,
      "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
      "Indexed": false,
      "InternalName": "PnPAlertStartDateTime",
      "JSLink": "clienttemplates.js",
      "PinnedToFiltersPane": false,
      "ReadOnlyField": false,
      "Required": false,
      "SchemaXml": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"1\"><Default>[today]</Default></Field>",
      "Scope": "/sites/portal",
      "Sealed": false,
      "ShowInFiltersPane": 0,
      "Sortable": true,
      "StaticName": "PnPAlertStartDateTime",
      "Title": "Start date-time",
      "FieldTypeKind": 4,
      "TypeAsString": "DateTime",
      "TypeDisplayName": "Date and Time",
      "TypeShortDescription": "Date and Time",
      "ValidationFormula": null,
      "ValidationMessage": null,
      "DateTimeCalendarType": 0,
      "DisplayFormat": 1,
      "FriendlyDisplayFormat": 1
    };
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/CreateFieldAsXml`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 52
          }
        })) {
        return Promise.resolve(response);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'AddToAllContentTypes, AddFieldToDefaultView, AddFieldCheckDisplayName' } });
    assert(loggerLogSpy.calledWith(response));
  });

  it('creates site column using XML with the AddToDefaultContentType, AddFieldInternalNameHint options', async () => {
    const response = {
      "AutoIndexed": false,
      "CanBeDeleted": true,
      "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
      "ClientSideComponentProperties": null,
      "CustomFormatter": null,
      "DefaultFormula": null,
      "DefaultValue": "[today]",
      "Description": "",
      "Direction": "none",
      "EnforceUniqueValues": false,
      "EntityPropertyName": "PnPAlertStartDateTime",
      "Filterable": true,
      "FromBaseType": false,
      "Group": "PnP Columns",
      "Hidden": false,
      "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
      "Indexed": false,
      "InternalName": "PnPAlertStartDateTime",
      "JSLink": "clienttemplates.js",
      "PinnedToFiltersPane": false,
      "ReadOnlyField": false,
      "Required": false,
      "SchemaXml": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"1\"><Default>[today]</Default></Field>",
      "Scope": "/sites/portal",
      "Sealed": false,
      "ShowInFiltersPane": 0,
      "Sortable": true,
      "StaticName": "PnPAlertStartDateTime",
      "Title": "Start date-time",
      "FieldTypeKind": 4,
      "TypeAsString": "DateTime",
      "TypeDisplayName": "Date and Time",
      "TypeShortDescription": "Date and Time",
      "ValidationFormula": null,
      "ValidationMessage": null,
      "DateTimeCalendarType": 0,
      "DisplayFormat": 1,
      "FriendlyDisplayFormat": 1
    };
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/CreateFieldAsXml`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 9
          }
        })) {
        return Promise.resolve(response);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'AddToDefaultContentType, AddFieldInternalNameHint' } });
    assert(loggerLogSpy.calledWith(response));
  });

  it('creates site column using XML with the AddToNoContentType option', async () => {
    const response = {
      "AutoIndexed": false,
      "CanBeDeleted": true,
      "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
      "ClientSideComponentProperties": null,
      "CustomFormatter": null,
      "DefaultFormula": null,
      "DefaultValue": "[today]",
      "Description": "",
      "Direction": "none",
      "EnforceUniqueValues": false,
      "EntityPropertyName": "PnPAlertStartDateTime",
      "Filterable": true,
      "FromBaseType": false,
      "Group": "PnP Columns",
      "Hidden": false,
      "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
      "Indexed": false,
      "InternalName": "PnPAlertStartDateTime",
      "JSLink": "clienttemplates.js",
      "PinnedToFiltersPane": false,
      "ReadOnlyField": false,
      "Required": false,
      "SchemaXml": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"1\"><Default>[today]</Default></Field>",
      "Scope": "/sites/portal",
      "Sealed": false,
      "ShowInFiltersPane": 0,
      "Sortable": true,
      "StaticName": "PnPAlertStartDateTime",
      "Title": "Start date-time",
      "FieldTypeKind": 4,
      "TypeAsString": "DateTime",
      "TypeDisplayName": "Date and Time",
      "TypeShortDescription": "Date and Time",
      "ValidationFormula": null,
      "ValidationMessage": null,
      "DateTimeCalendarType": 0,
      "DisplayFormat": 1,
      "FriendlyDisplayFormat": 1
    };
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/CreateFieldAsXml`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 2
          }
        })) {
        return Promise.resolve(response);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'AddToNoContentType' } });
    assert(loggerLogSpy.calledWith(response));
  });

  it('creates list column to list retrieved by ID using XML with the DefaultValue option', async () => {
    const output = {
      "AutoIndexed": false,
      "CanBeDeleted": true,
      "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
      "ClientSideComponentProperties": null,
      "CustomFormatter": null,
      "DefaultFormula": null,
      "DefaultValue": "[today]",
      "Description": "",
      "Direction": "none",
      "EnforceUniqueValues": false,
      "EntityPropertyName": "PnPAlertStartDateTime",
      "Filterable": true,
      "FromBaseType": false,
      "Group": "PnP Columns",
      "Hidden": false,
      "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
      "Indexed": false,
      "InternalName": "PnPAlertStartDateTime",
      "JSLink": "clienttemplates.js",
      "PinnedToFiltersPane": false,
      "ReadOnlyField": false,
      "Required": false,
      "SchemaXml": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"1\"><Default>[today]</Default></Field>",
      "Scope": "/sites/portal",
      "Sealed": false,
      "ShowInFiltersPane": 0,
      "Sortable": true,
      "StaticName": "PnPAlertStartDateTime",
      "Title": "Start date-time",
      "FieldTypeKind": 4,
      "TypeAsString": "DateTime",
      "TypeDisplayName": "Date and Time",
      "TypeShortDescription": "Date and Time",
      "ValidationFormula": null,
      "ValidationMessage": null,
      "DateTimeCalendarType": 0,
      "DisplayFormat": 1,
      "FriendlyDisplayFormat": 1
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/lists(guid'7118909e-4102-4473-a677-be6bf828245e')/fields/CreateFieldAsXml` &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 0
          }
        })) {
        return output;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: '7118909e-4102-4473-a677-be6bf828245e', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'DefaultValue' } });
    assert(loggerLogSpy.calledWith(output));
  });

  it('creates list column to list retrieved by list URL using XML with the DefaultValue option', async () => {
    const webUrl = 'https://contoso.sharepoint.com/sites/sales';
    const listUrl = '/sites/sales/Documents';
    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

    const output = {
      "AutoIndexed": false,
      "CanBeDeleted": true,
      "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
      "ClientSideComponentProperties": null,
      "CustomFormatter": null,
      "DefaultFormula": null,
      "DefaultValue": "[today]",
      "Description": "",
      "Direction": "none",
      "EnforceUniqueValues": false,
      "EntityPropertyName": "PnPAlertStartDateTime",
      "Filterable": true,
      "FromBaseType": false,
      "Group": "PnP Columns",
      "Hidden": false,
      "Id": "5ee2dd25-d941-455a-9bdb-7f2c54aed11b",
      "Indexed": false,
      "InternalName": "PnPAlertStartDateTime",
      "JSLink": "clienttemplates.js",
      "PinnedToFiltersPane": false,
      "ReadOnlyField": false,
      "Required": false,
      "SchemaXml": "<Field Type=\"DateTime\" DisplayName=\"Start date-time\" Required=\"FALSE\" EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" Format=\"DateTime\" Group=\"PnP Columns\" FriendlyDisplayFormat=\"Disabled\" ID=\"{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}\" SourceID=\"{4f118c69-66e0-497c-96ff-d7855ce0713d}\" StaticName=\"PnPAlertStartDateTime\" Name=\"PnPAlertStartDateTime\" Version=\"1\"><Default>[today]</Default></Field>",
      "Scope": "/sites/portal",
      "Sealed": false,
      "ShowInFiltersPane": 0,
      "Sortable": true,
      "StaticName": "PnPAlertStartDateTime",
      "Title": "Start date-time",
      "FieldTypeKind": 4,
      "TypeAsString": "DateTime",
      "TypeDisplayName": "Date and Time",
      "TypeShortDescription": "Date and Time",
      "ValidationFormula": null,
      "ValidationMessage": null,
      "DateTimeCalendarType": 0,
      "DisplayFormat": 1,
      "FriendlyDisplayFormat": 1
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/fields/CreateFieldAsXml` &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 0
          }
        })) {
        return output;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: webUrl, listUrl: listUrl, xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'DefaultValue' } });
    assert(loggerLogSpy.calledWith(output));
  });

  it('correctly handles a random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'AddToNoContentType' } } as any),
      new CommandError('An error has occurred'));
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

  it('fails validation if the specified site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'site.com', xml: '<Field />' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specified options is invalid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when option is set to DefaultValue', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'DefaultValue' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when option is set to AddToDefaultContentType', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddToDefaultContentType' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when option is set to AddToNoContentType', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddToNoContentType' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when option is set to AddToAllContentTypes', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddToAllContentTypes' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when option is set to AddFieldInternalNameHint', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldInternalNameHint' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when option is set to AddFieldToDefaultView', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldToDefaultView' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when option is set to AddFieldCheckDisplayName', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldCheckDisplayName' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when option is set to AddFieldCheckDisplayName and AddFieldToDefaultView', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldCheckDisplayName,AddFieldToDefaultView' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when option is set to AddFieldCheckDisplayName and AddFieldToDefaultView (with space)', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldCheckDisplayName, AddFieldToDefaultView' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if multiple list options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', listUrl: '/sites/sales/documents', listTitle: 'Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if list id is not a valid guid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', listId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
