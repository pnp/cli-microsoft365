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
import command, { options } from './list-defaultvalue-get.js';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
import { CommandError } from '../../../../Command.js';

describe(commands.LIST_DEFAULTVALUE_GET, () => {
  const siteUrl = 'https://contoso.sharepoint.com/sites/marketing';
  const listId = 'c090e594-3b8e-4f4d-9b9f-3e8e1f0b9f1a';
  const listTitle = 'Documents';
  const listUrl = '/sites/marketing/Shared Documents';
  const siteRelListUrl = '/Shared Documents';
  const folderUrl = '/sites/marketing/Shared Documents/Logos';
  const fieldName = 'DocumentType';

  const defaultColumnXml = `
  <MetadataDefaults>
    <a href="/sites/Marketing/Shared Documents">
      <DefaultValue FieldName="DocumentType">General</DefaultValue>
    </a>
    <a href="/sites/Marketing/Shared Documents/Logos">
      <DefaultValue FieldName="DocumentType">Logo</DefaultValue>
    </a>
  </MetadataDefaults>`;

  const defaultColumnValueRootLibrary = {
    fieldName: 'DocumentType',
    fieldValue: 'General',
    folderUrl: '/sites/Marketing/Shared Documents'
  };

  const defaultColumnValueFolder = {
    fieldName: 'DocumentType',
    fieldValue: 'Logo',
    folderUrl: '/sites/Marketing/Shared Documents/Logos'
  };

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

    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_DEFAULTVALUE_GET);
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
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId and listTitle are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, listTitle: listTitle, fieldName: fieldName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId and listUrl are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, listUrl: listUrl, fieldName: fieldName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if fieldName is not specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId });
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

  it('only outputs one single result', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists('${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listId: listId, fieldName: fieldName } });
    assert(loggerLogSpy.calledOnce);
  });

  it('correctly retrieves column default value for the specified field and list by listId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists('${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listId: listId, fieldName: fieldName, verbose: true } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValueRootLibrary);
  });

  it('correctly retrieves column default value for the specified field and list by listTitle', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists/GetByTitle('${listTitle}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listTitle: listTitle, fieldName: fieldName, verbose: true } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValueRootLibrary);
  });

  it('correctly retrieves column default value for the specified field and list by listUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName, verbose: true } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValueRootLibrary);
  });

  it('correctly retrieves column default value for the specified field and list by listUrl when using a site relative URL', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: siteRelListUrl, fieldName: fieldName, verbose: true } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValueRootLibrary);
  });

  it('correctly filters column default value for the specified folder', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: siteRelListUrl, fieldName: fieldName, folderUrl: folderUrl } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValueFolder);
  });

  it('correctly filters column default value for the specified folder with a site relative URL', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: siteRelListUrl, fieldName: fieldName, folderUrl: '/Shared Documents/LoGoS' } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValueFolder);
  });

  it('correctly logs error when field was not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
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

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: 'NonExistentField' } }),
      new CommandError("No default column value found for field 'NonExistentField'."));
  });

  it('correctly logs error when field default value not set for specified folder', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
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

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName, folderUrl: '/sites/marketing/Shared Documents/NonExistentFolder' } }),
      new CommandError("No default column value found for field 'DocumentType' in folder '/sites/marketing/Shared Documents/NonExistentFolder'."));
  });

  it('correctly handles when list has no default values set', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        throw { status: 404, error: { 'odata.error': { message: { value: 'The file does not exist.' } } } };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName } }),
      new CommandError(`No default column value found for field '${fieldName}'.`));
  });

  it('correctly handles error when retrieving column list', async () => {
    sinon.stub(request, 'get').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listId: listId } }),
      new CommandError('An error has occurred'));
  });

  it('correctly handles error when retrieving column default values', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      throw { error: { 'odata.error': { message: { value: 'An error has occurred' } } } };
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl } }),
      new CommandError('An error has occurred'));
  });

  it('correctly handles error when list is not found', async () => {
    sinon.stub(request, 'get').rejects({ status: 404 });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName } }),
      new CommandError(`List '${listUrl}' was not found.`));
  });

  it('fails command when list is not a document library', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 100 };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, fieldName: fieldName } }),
      new CommandError(`List '${listUrl}' is not a document library.`));
  });
});