import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command, { options } from './field-add.js';

describe(commands.FIELD_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()! as typeof options;
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FIELD_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates site column using XML with the default options', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/CreateFieldAsXml`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 0
          }
        })) {
        return {
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
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>' }) });
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
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Events')/fields/CreateFieldAsXml`) > -1 &&
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

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Events', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'DefaultValue' }) });
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
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/CreateFieldAsXml`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 52
          }
        })) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'AddToAllContentTypes, AddFieldToDefaultView, AddFieldCheckDisplayName' }) });
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
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/CreateFieldAsXml`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 9
          }
        })) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'AddToDefaultContentType, AddFieldInternalNameHint' }) });
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
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields/CreateFieldAsXml`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          parameters: {
            SchemaXml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>',
            Options: 2
          }
        })) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'AddToNoContentType' }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: '7118909e-4102-4473-a677-be6bf828245e', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'DefaultValue' }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, webUrl: webUrl, listUrl: listUrl, xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'DefaultValue' }) });
    assert(loggerLogSpy.calledWith(output));
  });

  it('correctly handles a random API error', async () => {
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

    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>', options: 'AddToNoContentType' }) } as any),
      new CommandError('An error has occurred'));
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'site.com', xml: '<Field />' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'site.com' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation if the specified options is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'invalid' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'invalid is not a valid value for the options argument. Allowed values are DefaultValue|AddToDefaultContentType|AddToNoContentType|AddToAllContentTypes|AddFieldInternalNameHint|AddFieldToDefaultView|AddFieldCheckDisplayName');
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when option is set to DefaultValue', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'DefaultValue' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when option is set to AddToDefaultContentType', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddToDefaultContentType' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when option is set to AddToNoContentType', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddToNoContentType' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when option is set to AddToAllContentTypes', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddToAllContentTypes' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when option is set to AddFieldInternalNameHint', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldInternalNameHint' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when option is set to AddFieldToDefaultView', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldToDefaultView' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when option is set to AddFieldCheckDisplayName', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldCheckDisplayName' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when option is set to AddFieldCheckDisplayName and AddFieldToDefaultView', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldCheckDisplayName,AddFieldToDefaultView' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when option is set to AddFieldCheckDisplayName and AddFieldToDefaultView (with space)', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', options: 'AddFieldCheckDisplayName, AddFieldToDefaultView' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if multiple list options are passed', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', listUrl: '/sites/sales/documents', listTitle: 'Documents' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'Specify either list id or title or list url, but not multiple');
  });

  it('fails validation if list id is not a valid guid', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', listId: 'foo' });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'foo in option listId is not a valid GUID');
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com/sites/sales', xml: '<Field />', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});
