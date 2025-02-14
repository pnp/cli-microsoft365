import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './list-defaultvalue-set.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
import { CommandError } from '../../../../Command.js';
import { urlUtil } from '../../../../utils/urlUtil.js';

describe(commands.LIST_DEFAULTVALUE_SET, () => {
  const siteUrl = 'https://contoso.sharepoint.com/sites/Marketing';
  const listId = 'c090e594-3b8e-4f4d-9b9f-3e8e1f0b9f1a';
  const listTitle = 'Documents';
  const listUrl = '/sites/Marketing/Shared Documents';
  const folderUrl = '/sites/Marketing/Shared Documents/Logos';
  const fieldName = 'DocumentType';
  const fieldValue = 'Logo';

  const defaultColumnXml = `<MetadataDefaults><a href="/sites/Marketing/Shared%20Documents"><DefaultValue FieldName="Countries">19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6</DefaultValue><DefaultValue FieldName="DocumentType">General</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Logos"><DefaultValue FieldName="Countries">20;#Canada|e3d25461-68ef-4070-8523-5ba439f6d4d5</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Templates"><DefaultValue FieldName="DocumentType">Template</DefaultValue></a></MetadataDefaults>`;

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    auth.connection.active = true;
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
      request.get,
      request.post,
      request.put
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_DEFAULTVALUE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', listId: listId });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId, listTitle and listUrl are not specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId and listTitle are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, listTitle: listTitle, fieldName: fieldName, fieldValue: fieldValue });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId and listUrl are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, listUrl: listUrl, fieldName: fieldName, fieldValue: fieldValue });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if fieldValue is an empty string', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, fieldValue: '', fieldName: fieldName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if folderUrl contains a # character', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, folderUrl: '/sites/marketing/Shared Documents/Logos#/Contoso', fieldName: fieldName, fieldValue: fieldValue });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if folderUrl contains a % character', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, folderUrl: '/sites/marketing/Shared Documents/Logos%/Contoso', fieldName: fieldName, fieldValue: fieldValue });
    assert.strictEqual(actual.success, false);
  });

  it('succeeds validation with folderUrl parameter', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, folderUrl: folderUrl, fieldName: fieldName, fieldValue: fieldValue });
    assert.strictEqual(actual.success, true);
  });

  it('succeeds validation without folderUrl parameter', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, fieldName: fieldName, fieldValue: fieldValue });
    assert.strictEqual(actual.success, true);
  });

  it('sets default column value for a field without generating output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return {
          BaseTemplate: 101,
          RootFolder: {
            ServerRelativeUrl: listUrl
          }
        };
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw `Invalid request: ${opts.url}`;
    });

    sinon.stub(request, 'put').resolves();

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldValue: fieldValue, fieldName: fieldName, verbose: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('updates an existing default column value', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return {
          BaseTemplate: 101,
          RootFolder: {
            ServerRelativeUrl: listUrl
          }
        };
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw `Invalid request: ${opts.url}`;
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return;
      }

      throw `Invalid request: ${opts.url}`;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listTitle: listTitle, fieldValue: fieldValue, fieldName: fieldName } });
    assert.deepStrictEqual(putStub.firstCall.args[0].data, `<MetadataDefaults><a href="/sites/Marketing/Shared%20Documents"><DefaultValue FieldName="Countries">19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6</DefaultValue><DefaultValue FieldName="DocumentType">Logo</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Logos"><DefaultValue FieldName="Countries">20;#Canada|e3d25461-68ef-4070-8523-5ba439f6d4d5</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Templates"><DefaultValue FieldName="DocumentType">Template</DefaultValue></a></MetadataDefaults>`);
  });

  it('adds a default column value to a folder that already has default values', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists('${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return {
          BaseTemplate: 101,
          RootFolder: {
            ServerRelativeUrl: listUrl
          }
        };
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(folderUrl)}')/ListItemAllFields?$select=FileRef`) {
        return {
          FileRef: folderUrl
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return;
      }

      throw `Invalid PUT request: ${opts.url}`;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listId: listId, fieldValue: fieldValue, fieldName: fieldName, folderUrl: folderUrl, verbose: true } });
    assert.deepStrictEqual(putStub.firstCall.args[0].data, `<MetadataDefaults><a href="/sites/Marketing/Shared%20Documents"><DefaultValue FieldName="Countries">19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6</DefaultValue><DefaultValue FieldName="DocumentType">General</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Logos"><DefaultValue FieldName="Countries">20;#Canada|e3d25461-68ef-4070-8523-5ba439f6d4d5</DefaultValue><DefaultValue FieldName="DocumentType">Logo</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Templates"><DefaultValue FieldName="DocumentType">Template</DefaultValue></a></MetadataDefaults>`);
  });

  it('adds a default column value to a list that has no default folder values', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return {
          BaseTemplate: 101,
          RootFolder: {
            ServerRelativeUrl: listUrl
          }
        };
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        throw { status: 404, error: { 'odata.error': { message: { value: 'The file does not exist.' } } } };
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(folderUrl)}')/ListItemAllFields?$select=FileRef`) {
        return {
          FileRef: folderUrl
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return;
      }

      throw `Invalid PUT request: ${opts.url}`;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms')}')/Files/Add(url='client_LocationBasedDefaults.html', overwrite=false)`) {
        return;
      }

      throw `Invalid POST request: ${opts.url}`;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldValue: fieldValue, fieldName: fieldName, folderUrl: folderUrl } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, '<MetadataDefaults />');
    assert.deepStrictEqual(putStub.firstCall.args[0].data, `<MetadataDefaults><a href="/sites/Marketing/Shared%20Documents/Logos"><DefaultValue FieldName="DocumentType">Logo</DefaultValue></a></MetadataDefaults>`);
  });

  it('adds a default column value correctly to a folder with an incorrect cased path', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists('${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return {
          BaseTemplate: 101,
          RootFolder: {
            ServerRelativeUrl: listUrl
          }
        };
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(siteUrl, folderUrl.toUpperCase()))}')/ListItemAllFields?$select=FileRef`) {
        return {
          FileRef: folderUrl
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return;
      }

      throw `Invalid PUT request: ${opts.url}`;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listId: listId, fieldValue: fieldValue, fieldName: fieldName, folderUrl: folderUrl.toUpperCase(), verbose: true } });
    assert.deepStrictEqual(putStub.firstCall.args[0].data, `<MetadataDefaults><a href="/sites/Marketing/Shared%20Documents"><DefaultValue FieldName="Countries">19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6</DefaultValue><DefaultValue FieldName="DocumentType">General</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Logos"><DefaultValue FieldName="Countries">20;#Canada|e3d25461-68ef-4070-8523-5ba439f6d4d5</DefaultValue><DefaultValue FieldName="DocumentType">Logo</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Templates"><DefaultValue FieldName="DocumentType">Template</DefaultValue></a></MetadataDefaults>`);
  });

  it('adds a default column value correctly when site relative url is used for folder', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists('${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return {
          BaseTemplate: 101,
          RootFolder: {
            ServerRelativeUrl: listUrl
          }
        };
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(folderUrl)}')/ListItemAllFields?$select=FileRef`) {
        return {
          FileRef: folderUrl
        };
      }

      throw `Invalid GET request: ${opts.url}`;
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return;
      }

      throw `Invalid PUT request: ${opts.url}`;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listId: listId, fieldValue: fieldValue, fieldName: fieldName, folderUrl: '/Shared Documents/Logos' } });
    assert.deepStrictEqual(putStub.firstCall.args[0].data, `<MetadataDefaults><a href="/sites/Marketing/Shared%20Documents"><DefaultValue FieldName="Countries">19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6</DefaultValue><DefaultValue FieldName="DocumentType">General</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Logos"><DefaultValue FieldName="Countries">20;#Canada|e3d25461-68ef-4070-8523-5ba439f6d4d5</DefaultValue><DefaultValue FieldName="DocumentType">Logo</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Templates"><DefaultValue FieldName="DocumentType">Template</DefaultValue></a></MetadataDefaults>`);
  });

  it('throws error when running the command on a non-document library', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists('${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return {
          BaseTemplate: 100,
          RootFolder: {
            ServerRelativeUrl: listUrl
          }
        };
      }

      throw `Invalid request: ${opts.url}`;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listId: listId, fieldName: fieldName, fieldValue: fieldValue } }),
      new CommandError('The specified list is not a document library.'));
  });

  it('throws error when list does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        throw { status: 404, error: { 'odata.error': { message: { value: 'The file does not exist.' } } } };
      }

      throw `Invalid request: ${opts.url}`;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName, fieldValue: fieldValue } }),
      new CommandError(`List '${listUrl}' was not found.`));
  });

  it('throws error when folder does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return {
          BaseTemplate: 101,
          RootFolder: {
            ServerRelativeUrl: listUrl
          }
        };
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(folderUrl)}')/ListItemAllFields?$select=FileRef`) {
        return {
          FileRef: null
        };
      }

      throw `Invalid request: ${opts.url}`;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName, fieldValue: fieldValue, folderUrl: folderUrl } }),
      new CommandError(`Folder '${folderUrl}' was not found.`));
  });

  it('throws error when error occurs when retrieving default column values file', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return {
          BaseTemplate: 101,
          RootFolder: {
            ServerRelativeUrl: listUrl
          }
        };
      }
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        throw { status: 401, error: { 'odata.error': { message: { value: 'You don\'t have permission to view this file.' } } } };
      }

      throw `Invalid request: ${opts.url}`;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName, fieldValue: fieldValue } }),
      new CommandError('You don\'t have permission to view this file.'));
  });
});