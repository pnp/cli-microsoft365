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
import command from './list-defaultvalue-remove.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
import { CommandError } from '../../../../Command.js';

describe(commands.LIST_DEFAULTVALUE_REMOVE, () => {
  const siteUrl = 'https://contoso.sharepoint.com/sites/Marketing';
  const listId = 'c090e594-3b8e-4f4d-9b9f-3e8e1f0b9f1a';
  const listTitle = 'Documents';
  const listUrl = '/sites/Marketing/Shared Documents';
  const folderUrl = '/sites/Marketing/Shared Documents/Logos';
  const siteRelativeFolderUrl = '/Shared Documents/Logos';
  const fieldName = 'DocumentType';
  const fieldValue = 'Logo';

  const defaultColumnXml = `<MetadataDefaults><a href="/sites/Marketing/Shared%20Documents"><DefaultValue FieldName="Countries">19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6</DefaultValue><DefaultValue FieldName="DocumentType">General</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Logos"><DefaultValue FieldName="DocumentType">Template</DefaultValue></a></MetadataDefaults>`;

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let confirmationPromptStub: sinon.SinonStub;

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
    confirmationPromptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.put,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_DEFAULTVALUE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'invalid', listId: listId, fieldName: fieldName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: 'invalid', fieldName: fieldName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId, listTitle and listUrl are not specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, fieldName: fieldName });
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

  it('succeeds validation if folderUrl is specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, fieldName: fieldName, folderUrl: folderUrl });
    assert.strictEqual(actual.success, true);
  });

  it('succeeds validation if folderUrl is not specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, fieldName: fieldName });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before removing default values', async () => {
    await command.action(logger, { options: { webUrl: siteUrl, listId: listId, fieldName: fieldName, folderUrl: folderUrl } });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts removing default values when prompt not confirmed', async () => {
    const putStub = sinon.stub(request, 'put').resolves();
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

      throw `Invalid GET request: ${opts.url}`;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName } });
    assert(putStub.notCalled);
  });

  it('logs no output when a default column value has been removed', async () => {
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

      throw `Invalid GET request: ${opts.url}`;
    });

    sinon.stub(request, 'put').resolves();

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName, force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly removes a default column value in the root of the list', async () => {
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

      throw `Invalid GET request: ${opts.url}`;
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return;
      }

      throw `Invalid PUT request: ${opts.url}`;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listTitle: listTitle, fieldName: fieldName, force: true, verbose: true } });
    assert.deepStrictEqual(putStub.firstCall.args[0].data, '<MetadataDefaults><a href="/sites/Marketing/Shared%20Documents"><DefaultValue FieldName="Countries">19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6</DefaultValue></a><a href="/sites/Marketing/Shared%20Documents/Logos"><DefaultValue FieldName="DocumentType">Template</DefaultValue></a></MetadataDefaults>');
  });

  it('correctly removes a default column value in on a sub-folder with site relative URL', async () => {
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

      throw `Invalid GET request: ${opts.url}`;
    });

    const putStub = sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return;
      }

      throw `Invalid PUT request: ${opts.url}`;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listId: listId, fieldName: fieldName, folderUrl: siteRelativeFolderUrl, force: true, verbose: true } });
    assert.deepStrictEqual(putStub.firstCall.args[0].data, '<MetadataDefaults><a href="/sites/Marketing/Shared%20Documents"><DefaultValue FieldName="Countries">19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6</DefaultValue><DefaultValue FieldName="DocumentType">General</DefaultValue></a></MetadataDefaults>');
  });

  it('correctly logs error when field was not found', async () => {
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

      throw `Invalid GET request: ${opts.url}`;
    });

    const putStub = sinon.stub(request, 'put').resolves();

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: 'Foo', force: true } }),
      new CommandError("Default column value 'Foo' was not found."));
    assert(putStub.notCalled);
  });

  it('throws error when list does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        throw { status: 404, error: { 'odata.error': { message: { value: 'The file does not exist.' } } } };
      }

      throw `Invalid GET request: ${opts.url}`;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName, force: true } }),
      new CommandError(`List '${listUrl}' was not found.`));
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

      throw `Invalid GET request: ${opts.url}`;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listId: listId, fieldName: fieldName, force: true } }),
      new CommandError('The specified list is not a document library.'));
  });

  it('correctly handles removing column default value for list that has no default values', async () => {
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

      throw `Invalid GET request: ${opts.url}`;
    });
    const putStub = sinon.stub(request, 'put').resolves();

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName, force: true } }),
      new CommandError(`Default column value '${fieldName}' was not found.`));
    assert(putStub.notCalled);
  });

  it('correctly handles error when unable to read column default values', async () => {
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

      throw `Invalid GET request: ${opts.url}`;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName, force: true } }),
      new CommandError('You don\'t have permission to view this file.'));
  });
});