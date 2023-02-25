import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import request from '../../../../request';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import Command, { CommandError } from '../../../../Command';
import commands from '../../commands';
import { Logger } from '../../../../cli/Logger';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
const command: Command = require('./listitem-batch-set');

describe(commands.LISTITEM_BATCH_SET, () => {
  const filePath = 'C:\\Path\\To\\CSV\\CsvFile.csv';
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const idColumn = 'Id';
  const listId = 'f2978459-4e2a-4307-b57c-0c90eb4e5d6a';
  const listTitle = 'Random List';
  const listUrl = '/sites/project-x/lists/random-list';
  const csvContentHeaders = `Id,ContentType,Title,SingleChoiceField,MultiChoiceField,SingleMetadataField,MultiMetadataField,SinglePeopleField,MultiPeopleField,CustomHyperlink,NumberField,LookupField,LookupFieldMulti`;
  const csvContentLine = `10,Item,Title A,Choice 1,Choice 1;#Choice 2,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760;,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760;Finance|f994a4ac-cf34-448e-a22c-2b35fd9bbffa;,[{'Key':'i:0#.f|membership|markh@contoso.com'}],"[{'Key':'i:0#.f|membership|markh@contoso.com'},{'Key':'i:0#.f|membership|adamb@contoso.com'}]","https://bing.com, URL",5,1,1;2`;
  const csvContent = `${csvContentHeaders}\n${csvContentLine}`;
  const fieldsResponse = [{ 'InternalName': 'ContentType', 'TypeAsString': 'Computed' }, { 'InternalName': 'Title', 'TypeAsString': 'Text' }, { 'InternalName': 'SingleChoiceField', 'TypeAsString': 'Choice' }, { 'InternalName': 'MultiChoiceField', 'TypeAsString': 'MultiChoice' }, { 'InternalName': 'SingleMetadataField', 'TypeAsString': 'TaxonomyFieldType' }, { 'InternalName': 'MultiMetadataField', 'TypeAsString': 'TaxonomyFieldTypeMulti' }, { 'InternalName': 'SinglePeopleField', 'TypeAsString': 'User' }, { 'InternalName': 'MultiPeopleField', 'TypeAsString': 'UserMulti' }, { 'InternalName': 'CustomHyperlink', 'TypeAsString': 'URL' }, { 'InternalName': 'NumberField', 'TypeAsString': 'Number' }, { 'InternalName': 'LookupField', 'TypeAsString': 'Lookup' }, { 'InternalName': 'LookupFieldMulti', 'TypeAsString': 'LookupMulti' }];
  const filterFields = ["InternalName eq 'ContentType'", "InternalName eq 'Title'", "InternalName eq 'SingleChoiceField'", "InternalName eq 'MultiChoiceField'", "InternalName eq 'SingleMetadataField'", "InternalName eq 'MultiMetadataField'", "InternalName eq 'SinglePeopleField'", "InternalName eq 'MultiPeopleField'", "InternalName eq 'CustomHyperlink'", "InternalName eq 'NumberField'", "InternalName eq 'LookupField'", "InternalName eq 'LookupFieldMulti'"];
  const batchCsomResponse = [{ 'SchemaVersion': '15.0.0.0', 'LibraryVersion': '16.0.23408.12001', 'ErrorInfo': null, 'TraceCorrelationId': '9c7d99a0-9005-6000-4c2b-7d8f8a647714' }];
  const listResponse = { Id: listId };

  let commandInfo: CommandInfo;
  let log: any[];
  let logger: Logger;

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
    sinon.stub(spo, 'getRequestDigest').callsFake(async () => ({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: webUrl
    }));
    sinon.stub(spo, 'getCurrentWebIdentity').callsFake(async () => ({
      'objectIdentity': '04e9249b-1edd-40da-9ec9-c3f19b2db1bd|25a633e6-3138-49c0-8be8-8bd3260a0431:site:339fb67e-4573-4eee-91b8-7e4fdb1a38d7:web:2c82aec1-21d2-4a1a-ad95-15bb7a4b66aa',
      'serverRelativeUrl': '/sites/project-x'
    }));
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      odata.getAllItems,
      request.get,
      request.post,
      spo.getCurrentWebIdentity,
      spo.getRequestDigest
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
    assert.strictEqual(command.name, commands.LISTITEM_BATCH_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates single item in batch to a sharepoint list retrieved by listUrl', async () => {
    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);

    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')?$select=Id`) {
        return listResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields?$select=InternalName,TypeAsString&$filter=${filterFields.join(' or ')}`) {
        return fieldsResponse;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify(batchCsomResponse);
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listUrl: listUrl, idColumn: idColumn, systemUpdate: true, verbose: true } } as any);
    assert(postStub.called);
  });

  it('system updates single item in batch to a sharepoint list retrieved by id', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields?$select=InternalName,TypeAsString&$filter=${filterFields.join(' or ')}`) {
        return fieldsResponse;
      }

      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify(batchCsomResponse);
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listId: listId, idColumn: idColumn, systemUpdate: true, verbose: true } } as any);
    assert(postStub.called);
  });

  it('updates items in multiple batches to a sharepoint list retrieved by title', async () => {
    let amountOfExecutions = 0;
    let csvContent150Items = csvContent;
    for (let i = 1; i < 150; i++) {
      csvContent150Items += `\n${csvContentLine}`;
    }
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent150Items);
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields?$select=InternalName,TypeAsString&$filter=${filterFields.join(' or ')}`) {
        return fieldsResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')?$select=Id`) {
        return listResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`) {
        amountOfExecutions++;
        return JSON.stringify(batchCsomResponse);
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listId: listId, idColumn: idColumn, verbose: true } } as any);
    assert.strictEqual(amountOfExecutions, 3);
  });

  it('throws an error when a wrong value is entered (text instead of number)', async () => {
    const csvContentLine = `10,Item,Title A,Choice 1,Choice 1;#Choice 2,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760;,Engineering|4a3cc5f3-a4a6-433e-a07a-746978ff1760;Finance|f994a4ac-cf34-448e-a22c-2b35fd9bbffa;,[{'Key':'i:0#.f|membership|markh@contoso.com'}],"[{'Key':'i:0#.f|membership|markh@contoso.com'},{'Key':'i:0#.f|membership|adamb@contoso.com'}]","https://bing.com, URL",'TEXT',1,1;2`;
    const csvContent = `${csvContentHeaders}\n${csvContentLine}`;
    const batchCsomResponseError = [{ 'SchemaVersion': '15.0.0.0', 'LibraryVersion': '16.0.23408.12001', 'ErrorInfo': { 'ErrorMessage': 'Only numbers can go here.', 'ErrorValue': '362,NumberField', 'TraceCorrelationId': '4d7f99a0-3064-6000-40b7-61b9fc6fcd53', 'ErrorCode': -2130575155, 'ErrorTypeName': 'Microsoft.SharePoint.SPFieldValueException' }, 'TraceCorrelationId': '4d7f99a0-3064-6000-40b7-61b9fc6fcd53' }];
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields?$select=InternalName,TypeAsString&$filter=${filterFields.join(' or ')}`) {
        return fieldsResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts: any) => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return JSON.stringify(batchCsomResponseError);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listId: listId, idColumn: idColumn, verbose: true } } as any), new CommandError(`${batchCsomResponseError[0].ErrorInfo.ErrorMessage} - ${batchCsomResponseError[0].ErrorInfo.ErrorValue}`));
  });

  it('throws an error when field specified in the csv does not exist on the list', async () => {
    const fieldsThatDontExist = ['NonExistingColumn1', 'NonExistingColumn2'];
    const errorMessage = `Following fields specified in the csv do not exist on the list: ${fieldsThatDontExist.join(', ')}`;
    const csvContentHeadersError = csvContentHeaders + `,${fieldsThatDontExist.join(',')}`;
    const csvContentLineError = csvContentLine + ',Value 1,Value2';
    const csvContentError = `${csvContentHeadersError}\n${csvContentLineError}`;
    const jsonContent: any[] = formatting.parseCsvToJson(csvContentError);

    const objectKeys = Object.keys(jsonContent[0]);
    const index = objectKeys.indexOf(idColumn, 0);
    if (index > -1) {
      objectKeys.splice(index, 1);
    }

    const filterFields: string[] = [];
    objectKeys.map(objectKey => {
      filterFields.push(`InternalName eq '${objectKey}'`);
    });

    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContentError);
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/fields?$select=InternalName,TypeAsString&$filter=${filterFields.join(' or ')}`) {
        return fieldsResponse;
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listId: listId, idColumn: idColumn, verbose: true } } as any), new CommandError(errorMessage));
  });

  it('throws an error when list by url is not found', async () => {
    const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
    const errorMessage = `File Not Found.`;
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')?$select=Id`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listUrl: listUrl, idColumn: idColumn, verbose: true } } as any), new CommandError(errorMessage));
  });

  it('throws an error when list by title is not found', async () => {
    const errorMessage = `List '${listTitle}' does not exist at site with URL 'https://mathijsdev2.sharepoint.com'.`;
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')?$select=Id`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listTitle: listTitle, idColumn: idColumn, verbose: true } } as any), new CommandError(errorMessage));
  });

  it('throws an error when specified idColumn does not exist in csv', async () => {
    const tempIdColumn = 'id';
    const errorMessage = `The specified value for idColumn does not exist in the array. Specified idColumn is '${tempIdColumn || 'ID'}'. Please specify the correct value.`;
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listUrl: listUrl, idColumn: tempIdColumn, verbose: true } } as any), new CommandError(errorMessage));
  });

  it('throws an error when idColumn is not specified and ID does not exist in csv', async () => {
    const errorMessage = `The specified value for idColumn does not exist in the array. Specified idColumn is 'ID'. Please specify the correct value.`;
    sinon.stub(fs, 'readFileSync').callsFake(_ => csvContent);

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, filePath: filePath, listUrl: listUrl, verbose: true } } as any), new CommandError(errorMessage));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', filePath: filePath, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId option is not a valid GUID', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if csv file does not exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => false);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if filePath exists, listId is a valid guid and idColumn is a valid idColumn', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    const actual = await command.validate({ options: { webUrl: webUrl, filePath: filePath, listId: listId, idColumn: idColumn } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
