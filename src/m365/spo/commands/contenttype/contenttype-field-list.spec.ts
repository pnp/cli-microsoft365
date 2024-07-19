import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './contenttype-field-list.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { CommandError } from '../../../../Command.js';

describe(commands.CONTENTTYPE_FIELD_LIST, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const listId = '0c8a1a3b-1685-4b4c-9b4c-3c1c8b2f7f8f';
  const listUrl = 'Lists/MyList';
  const listTitle = 'MyList';
  const contentTypeId = '0x0100A2B3CD4E5F6A7B8C9D0E1F2A3B4C5D6E7F8';
  const contentTypeName = 'My Content Type';
  const properties = 'Id,InternalName,Group';
  const fieldResponse = [{
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
    "EntityPropertyName": "Modified_x0020_By",
    "Filterable": true,
    "FromBaseType": false,
    "Group": "_Hidden",
    "Hidden": false,
    "Id": "822c78e3-1ea9-4943-b449-57863ad33ca9",
    "Indexed": false,
    "IndexStatus": 0,
    "InternalName": "Modified_x0020_By",
    "IsModern": false,
    "JSLink": "clienttemplates.js",
    "PinnedToFiltersPane": false,
    "ReadOnlyField": true,
    "Required": false,
    "SchemaXml": "<Field ID=\"{822c78e3-1ea9-4943-b449-57863ad33ca9}\" Name=\"Modified_x0020_By\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Modified_x0020_By\" Group=\"_Hidden\" ReadOnly=\"TRUE\" Hidden=\"FALSE\" Type=\"Text\" DisplayName=\"Document Modified By\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\" />",
    "Scope": "/",
    "Sealed": false,
    "ShowInFiltersPane": 0,
    "Sortable": true,
    "StaticName": "Modified_x0020_By",
    "Title": "Document Modified By",
    "FieldTypeKind": 2,
    "TypeAsString": "Text",
    "TypeDisplayName": "Single line of text",
    "TypeShortDescription": "Single line of text",
    "ValidationFormula": null,
    "ValidationMessage": null,
    "MaxLength": 255
  },
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
    "EntityPropertyName": "Created_x0020_By",
    "Filterable": true,
    "FromBaseType": false,
    "Group": "_Hidden",
    "Hidden": false,
    "Id": "4dd7e525-8d6b-4cb4-9d3e-44ee25f973eb",
    "Indexed": false,
    "IndexStatus": 0,
    "InternalName": "Created_x0020_By",
    "IsModern": false,
    "JSLink": "clienttemplates.js",
    "PinnedToFiltersPane": false,
    "ReadOnlyField": true,
    "Required": false,
    "SchemaXml": "<Field ID=\"{4dd7e525-8d6b-4cb4-9d3e-44ee25f973eb}\" Name=\"Created_x0020_By\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Created_x0020_By\" Group=\"_Hidden\" ReadOnly=\"TRUE\" Hidden=\"FALSE\" Type=\"Text\" DisplayName=\"Document Created By\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\"></Field>",
    "Scope": "/",
    "Sealed": false,
    "ShowInFiltersPane": 0,
    "Sortable": true,
    "StaticName": "Created_x0020_By",
    "Title": "Document Created By",
    "FieldTypeKind": 2,
    "TypeAsString": "Text",
    "TypeDisplayName": "Single line of text",
    "TypeShortDescription": "Single line of text",
    "ValidationFormula": null,
    "ValidationMessage": null,
    "MaxLength": 255
  }];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTENTTYPE_FIELD_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', contentTypeId: contentTypeId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId is not a GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, contentTypeId: contentTypeId, listId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when webUrl, contentTypeId and listId are specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, contentTypeId: contentTypeId, listId: listId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves fields of a specific content type by name from a list identified by its title', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/contentTypes?$filter=Name eq '${formatting.encodeQueryParameter(contentTypeName)}'&$select=StringId`) {
        return [{ StringId: contentTypeId }];
      }

      if (url === `${webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/contentTypes('${contentTypeId}')/fields`) {
        return fieldResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, contentTypeName: contentTypeName, listTitle: listTitle, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(fieldResponse));
  });

  it('retrieves fields of a specific content type by name from a list identified by its url', async () => {
    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/contentTypes?$filter=Name eq '${formatting.encodeQueryParameter(contentTypeName)}'&$select=StringId`) {
        return [{ StringId: contentTypeId }];
      }

      if (url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/contentTypes('${contentTypeId}')/fields`) {
        return fieldResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, contentTypeName: contentTypeName, listUrl: listUrl, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(fieldResponse));
  });

  it('retrieves fields of a specific content type by id from a list identified by id', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/contentTypes('${contentTypeId}')/fields`) {
        return fieldResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, contentTypeId: contentTypeId, listId: listId, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(fieldResponse));
  });

  it('retrieves fields of a specific content type by id from the root web and adds properties to select statement', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/contentTypes('${contentTypeId}')/fields?$select=${properties}`) {
        return fieldResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, contentTypeId: contentTypeId, properties: properties, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(fieldResponse));
  });

  it('handles failure when content type specified by name is not found', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/contentTypes?$filter=Name eq '${formatting.encodeQueryParameter(contentTypeName)}'&$select=StringId`) {
        return [];
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, contentTypeName: contentTypeName, verbose: true } } as any),
      new CommandError(`Content type with name ${contentTypeName} not found.`));
  });

  it('handles failure when content type specified by id is not found', async () => {
    const error = {
      error: {
        code: '-2147024809, System.ArgumentException',
        message: 'Value does not fall within the expected range.'
      }
    };

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/contentTypes('${contentTypeId}')/fields?$select=${properties}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, contentTypeId: contentTypeId, properties: properties, verbose: true } } as any),
      new CommandError(error.error.message));
  });
});