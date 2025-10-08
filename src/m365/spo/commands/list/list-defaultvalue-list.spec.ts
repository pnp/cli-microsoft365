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
import command from './list-defaultvalue-list.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
import { CommandError } from '../../../../Command.js';

describe(commands.LIST_DEFAULTVALUE_LIST, () => {
  const siteUrl = 'https://contoso.sharepoint.com/sites/marketing';
  const listId = 'c090e594-3b8e-4f4d-9b9f-3e8e1f0b9f1a';
  const listTitle = 'Documents';
  const listUrl = '/sites/marketing/Shared Documents';
  const siteRelListUrl = '/Shared Documents';

  const defaultColumnXml = `
  <MetadataDefaults>
    <a href="/sites/Marketing/Shared Documents">
      <DefaultValue FieldName="Countries">19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6</DefaultValue>
      <DefaultValue FieldName="DocumentType">General</DefaultValue>
    </a>
    <a href="/sites/Marketing/Shared Documents/Logos">
      <DefaultValue FieldName="Countries">20;#Canada|e3d25461-68ef-4070-8523-5ba439f6d4d5</DefaultValue>
      <DefaultValue FieldName="DocumentType">Logo</DefaultValue>
    </a>
    <a href="/sites/Marketing/Shared Documents/Templates">
      <DefaultValue FieldName="DocumentType">Template</DefaultValue>
    </a>
  </MetadataDefaults>`;

  const defaultColumnValues = [
    {
      fieldName: 'Countries',
      fieldValue: '19;#Belgium|442affc2-7fab-4f33-9590-330403a579c2;#18;#Croatia|59f1ab85-235b-4cf8-b669-4373cc9393c6',
      folderUrl: '/sites/Marketing/Shared Documents'
    },
    {
      fieldName: 'DocumentType',
      fieldValue: 'General',
      folderUrl: '/sites/Marketing/Shared Documents'
    },
    {
      fieldName: 'Countries',
      fieldValue: '20;#Canada|e3d25461-68ef-4070-8523-5ba439f6d4d5',
      folderUrl: '/sites/Marketing/Shared Documents/Logos'
    },
    {
      fieldName: 'DocumentType',
      fieldValue: 'Logo',
      folderUrl: '/sites/Marketing/Shared Documents/Logos'
    },
    {
      fieldName: 'DocumentType',
      fieldValue: 'Template',
      folderUrl: '/sites/Marketing/Shared Documents/Templates'
    }
  ];

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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_DEFAULTVALUE_LIST);
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
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, listTitle: listTitle });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId and listUrl are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: siteUrl, listId: listId, listUrl: listUrl });
    assert.strictEqual(actual.success, false);
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

    await command.action(logger, { options: { webUrl: siteUrl, listId: listId } });
    assert(loggerLogSpy.calledOnce);
  });

  it('correctly retrieves column default values for the specified list by listId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists('${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listId: listId, verbose: true } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValues);
  });

  it('correctly retrieves column default values for the specified list by listTitle', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/Lists/GetByTitle('${listTitle}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listTitle: listTitle, verbose: true } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValues);
  });

  it('correctly retrieves column default values for the specified list by listUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl, verbose: true } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValues);
  });

  it('correctly retrieves column default values for the specified list by listUrl when using a site relative URL', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: siteRelListUrl, verbose: true } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValues);
  });

  it('correctly filters column default values for the specified folder', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: siteRelListUrl, folderUrl: '/sites/Marketing/Shared Documents/Logos' } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValues.filter(d => d.folderUrl.toLowerCase() === '/sites/marketing/shared documents/logos'));
  });

  it('correctly filters column default values for the specified folder with a site relative URL', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        return defaultColumnXml;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: siteRelListUrl, folderUrl: '/Shared Documents/TeMpLaTeS' } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], defaultColumnValues.filter(d => d.folderUrl.toLowerCase() === '/sites/marketing/shared documents/templates'));
  });

  it('correctly retrieves column default values for list without any default values set', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 101 };
      }

      if (opts.url === `${siteUrl}/_api/Web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(listUrl + '/Forms/client_LocationBasedDefaults.html')}')/$value`) {
        throw { status: 404, error: { 'odata.error': { message: { value: 'The file does not exist.' } } } };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl } });
    assert.deepStrictEqual(loggerLogSpy.firstCall.args[0], []);
  });

  it('correctly handles error when list is not found', async () => {
    sinon.stub(request, 'get').rejects({ status: 404 });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl } }),
      new CommandError(`List '${listUrl}' was not found.`));
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

  it('fails command when list is not a document library', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/Web/GetList('${listUrl}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl,BaseTemplate`) {
        return { RootFolder: { ServerRelativeUrl: listUrl }, BaseTemplate: 100 };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { webUrl: siteUrl, listUrl: listUrl } }),
      new CommandError(`List '${listUrl}' is not a document library.`));
  });
});