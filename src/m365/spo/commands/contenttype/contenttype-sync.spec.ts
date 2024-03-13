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
import command from './contenttype-sync.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import { CommandError } from '../../../../Command.js';

describe(commands.CONTENTTYPE_SYNC, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const contentTypeId = '0x0101000728167CD9C94899925BA69C4A5F9F3A';
  const contentTypeName = 'Dummy 8';
  const listId = 'd4552f22-5fb0-4df3-b216-309264237d2b';
  const listTitle = 'Documents';
  const listUrl = '/sites/project-x/Shared Documents';
  const graphBaseUrl = 'https://graph.microsoft.com/v1.0/sites/';
  const siteId = 'contoso.sharepoint.com,777794dc-43c9-4f36-88bd-a42721c75304,95c1171f-7b40-46bb-8679-3b5263dc019a';

  const syncResponse = {
    '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#contentType',
    '@odata.type': '#microsoft.graph.contentType',
    '@odata.etag': '\'2\'',
    id: '0x010100B01336624A574D47BE29121892EA4D98',
    isBuiltIn: false,
    description: '',
    group: 'Document Content Types',
    hidden: false,
    name: 'Dummy 8',
    parentId: '0x0101',
    readOnly: true,
    sealed: false,
    base: {
      id: '0x0101',
      description: 'Create a new document.',
      group: 'Document Content Types',
      hidden: false,
      name: 'Document',
      readOnly: false,
      sealed: false
    }
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getSiteId').resolves(siteId);
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
      request.post,
      odata.getAllItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTENTTYPE_SYNC);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('succesfully sync a content type to the site by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${graphBaseUrl}${siteId}/contenttypes/addCopyFromContentTypeHub` && opts.data.contentTypeId === contentTypeId) {
        return syncResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: contentTypeId, verbose: true } } as any);
    assert(loggerLogSpy.calledWith(syncResponse));
  });

  it('succesfully sync a content type to the site by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${graphBaseUrl}${siteId}/contenttypes/addCopyFromContentTypeHub` && opts.data.contentTypeId === contentTypeId) {
        return syncResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `${graphBaseUrl}${siteId}/contenttypes?$filter=name eq '${contentTypeName}'&$select=id,name`) {
        return [{ id: contentTypeId }];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, name: contentTypeName, verbose: true } } as any);
    assert(loggerLogSpy.calledWith(syncResponse));
  });

  it('succesfully sync a content type to a list by listId', async () => {
    const url = webUrl.split('/sites/')[0];
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${graphBaseUrl}${new URL(url).host}/lists/${listId}/contenttypes/addCopyFromContentTypeHub` && opts.data.contentTypeId === contentTypeId) {
        return syncResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `${graphBaseUrl}${siteId}/contenttypes?$filter=name eq '${contentTypeName}'&$select=id,name`) {
        return [{ id: contentTypeId }];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: url, name: contentTypeName, listId: listId, verbose: true } } as any);
    assert(loggerLogSpy.calledWith(syncResponse));
  });

  it('succesfully sync a content type to a list by listTitle', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${graphBaseUrl}${siteId}/lists/${listTitle}/contenttypes/addCopyFromContentTypeHub` && opts.data.contentTypeId === contentTypeId) {
        return syncResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: contentTypeId, listTitle: listTitle, verbose: true } } as any);
    assert(loggerLogSpy.calledWith(syncResponse));
  });

  it('succesfully sync a content type to a list by listUrl', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${graphBaseUrl}${siteId}/lists/${listId}/contenttypes/addCopyFromContentTypeHub` && opts.data.contentTypeId === contentTypeId) {
        return syncResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$select=id`) {
        return { Id: listId };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, id: contentTypeId, listUrl: listUrl, verbose: true } } as any);
    assert(loggerLogSpy.calledWith(syncResponse));
  });

  it('correctly handles contentType not found in the hub', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `${graphBaseUrl}${siteId}/contenttypes?$filter=name eq '${contentTypeName}'&$select=id,name`) {
        return [];
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, name: contentTypeName, verbose: true } } as any),
      new CommandError(`Content type with name ${contentTypeName} not found.`));
  });

  it('correctly handles error that occurs when content type not published yet in the hub', async () => {
    const error = {
      error: {
        code: 'invalidRequest',
        message: 'No published content type available',
        innerError: {
          date: '2024-03-13T15:40:51',
          'request-id': 'ed2d9ed7-fd83-4c25-9d74-b2c6845a8ede',
          'client-request-id': 'ed2d9ed7-fd83-4c25-9d74-b2c6845a8ede'
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${graphBaseUrl}${siteId}/contenttypes/addCopyFromContentTypeHub` && opts.data.contentTypeId === contentTypeId) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: contentTypeId, verbose: true } } as any),
      new CommandError(error.error.message));
  });

  it('correctly handles error that occurs when content type requires a feature that has not yet been activated on the site', async () => {
    const error = {
      error: {
        code: 'generalException',
        message: `Content type '${contentTypeName}' cannot be published to this site because feature '43f41342-1a37-4372-8ca0-b44d881e4434' is not enabled.`,
        innerError: {
          date: '2024-03-13T15:42:57',
          'request-id': 'cbc15894-b699-4dbf-bf50-8ce83d4180d7',
          'client-request-id': 'cbc15894-b699-4dbf-bf50-8ce83d4180d7'
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${graphBaseUrl}${siteId}/contenttypes/addCopyFromContentTypeHub` && opts.data.contentTypeId === contentTypeId) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: contentTypeId, verbose: true } } as any),
      new CommandError(error.error.message));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', id: contentTypeId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: contentTypeId, listId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when webUrl and id are specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: contentTypeId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when webUrl, id and listId are specified and listId is a valid guid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: contentTypeId, listId: listId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});