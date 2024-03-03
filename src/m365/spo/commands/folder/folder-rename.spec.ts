import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';

import command from './folder-rename.js';

describe(commands.FOLDER_RENAME, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const folderRelSiteUrl = '/Shared Documents/Folder1';
  const folderRelServerUrl = '/sites/project-x/Shared Documents/Folder1';
  const newFolderName = 'New name';

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_RENAME);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('renames folder correctly by using server relative URL', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderRelServerUrl)}')/ListItemAllFields`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options:
      {
        verbose: true,
        webUrl: webUrl,
        url: folderRelServerUrl,
        name: newFolderName
      }
    });
    assert(patchStub.called);
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { FileLeafRef: newFolderName, Title: newFolderName });
  });

  it('renames folder correctly by using site relative URL', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderRelServerUrl)}')/ListItemAllFields`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, {
      options:
      {
        webUrl: webUrl,
        url: folderRelSiteUrl,
        name: newFolderName
      }
    });
    assert(patchStub.called);
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { FileLeafRef: newFolderName, Title: newFolderName });
  });

  it('handles API error correctly', async () => {
    sinon.stub(request, 'patch').resolves({ 'odata.null': true });

    await assert.rejects(command.action(logger, {
      options:
      {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents/Folder1',
        recycle: true
      }
    } as any), new CommandError('Folder not found.'));
  });

  it('fails validation if the webUrl option is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'abc', url: '/Shared Documents/test', name: 'abc' } }, commandInfo);
    assert.strictEqual(actual, "abc is not a valid SharePoint Online site URL");
  });

  it('passes validation when the url option specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com',
        url: '/Shared Documents/test',
        name: 'abc'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
