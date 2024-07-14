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
import command from './folder-set.js';

describe(commands.FOLDER_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const folderRelSiteUrl = '/Shared Documents/Folder1';
  const folderRelServerUrl = '/sites/project-x/Shared Documents/Folder1';
  const newFolderName = 'New name';
  const colorName = 'darkRed';
  const colorNumber = '1';

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
      request.patch,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_SET);
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
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, { FileLeafRef: newFolderName, Title: newFolderName });
  });

  it('changes color of folder by using a number', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/foldercoloring/stampcolor(DecodedUrl='${formatting.encodeQueryParameter(folderRelServerUrl)}')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: webUrl, url: folderRelSiteUrl, color: colorNumber } });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      coloringInformation: {
        ColorHex: colorNumber
      },
      newName: undefined
    });
  });

  it('changes color of folder by using a color name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/foldercoloring/stampcolor(DecodedUrl='${formatting.encodeQueryParameter(folderRelServerUrl)}')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: webUrl, url: folderRelSiteUrl, color: colorName } });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      coloringInformation: {
        ColorHex: colorNumber
      },
      newName: undefined
    });
  });

  it('changes both color and name of folder by using a color name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/foldercoloring/renamefolder(DecodedUrl='${formatting.encodeQueryParameter(folderRelServerUrl)}')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: webUrl, url: folderRelSiteUrl, color: colorName, name: newFolderName } });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      coloringInformation: {
        ColorHex: colorNumber
      },
      newName: newFolderName
    });
  });

  it('changes both color and name of folder by using a color number', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/foldercoloring/renamefolder(DecodedUrl='${formatting.encodeQueryParameter(folderRelServerUrl)}')`) {
        return;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, { options: { webUrl: webUrl, url: folderRelSiteUrl, color: colorNumber, name: newFolderName } });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      coloringInformation: {
        ColorHex: colorNumber
      },
      newName: newFolderName
    });
  });

  it('handles error when something errors when trying to change color of folder', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-2147213181, Microsoft.SharePoint.SPException',
          message: {
            lang: 'en-US',
            value: 'Failed to stamp the folder'
          }
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/foldercoloring/stampcolor(DecodedUrl='${formatting.encodeQueryParameter(folderRelServerUrl)}')`) {
        throw error;
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: webUrl, url: folderRelServerUrl, color: colorName } } as any)
      , new CommandError(error.error['odata.error'].message.value));
  });

  it('handles API error correctly', async () => {
    sinon.stub(request, 'patch').resolves({ 'odata.null': true });

    await assert.rejects(command.action(logger, {
      options:
      {
        debug: true,
        webUrl: webUrl,
        url: folderRelServerUrl,
        name: newFolderName
      }
    } as any), new CommandError('Folder not found.'));
  });

  it('fails validation if the webUrl option is not valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'abc', url: folderRelSiteUrl, name: newFolderName } }, commandInfo);
    assert.strictEqual(actual, "'abc' is not a valid SharePoint Online site URL.");
  });

  it('passes validation when the url option specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: webUrl,
        url: folderRelServerUrl,
        name: newFolderName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if color is passed as string and color is not a valid color', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, url: folderRelSiteUrl, color: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if color is passed as string and color is a valid color', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, url: folderRelSiteUrl, color: colorName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if color is passed as number and color is a valid color', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, url: folderRelSiteUrl, color: colorNumber } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if neither name nor color is specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, url: folderRelSiteUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
