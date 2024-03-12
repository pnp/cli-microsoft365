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
import commands from '../../commands.js';
import command from './folder-add.js';

describe(commands.FOLDER_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let stubPostResponses: any;

  const addResponse = { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "abc", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/abc", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" };

  const webUrl = 'https://contoso.sharepoint.com';
  const parentFolder = '/Shared Documents';
  const folderName = 'My Folder';
  const colorName = 'darkRed';
  const colorNumber = '1';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;

    stubPostResponses = (addResp: any = null) => {
      return sinon.stub(request, 'post').callsFake(async (opts) => {
        if ((opts.url as string).indexOf('/_api/web/folders') > -1) {
          if (addResp) {
            throw addResp;
          }
          else {
            return addResponse;
          }
        }

        throw 'Invalid request';
      });
    };
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FOLDER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should correctly handle folder add reject request', async () => {
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
    stubPostResponses(error);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        name: 'abc'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('should correctly handle folder add success request', async () => {
    stubPostResponses();

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        name: 'abc'
      }
    });
    assert(loggerLogSpy.lastCall.calledWith({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "abc", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/abc", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" }));
  });

  it('should correctly pass params to request', async () => {
    const request: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        name: 'abc'
      }
    });
    assert(request.calledWith({
      url: `https://contoso.sharepoint.com/_api/web/folders/addUsingPath(decodedUrl='${formatting.encodeQueryParameter('/Shared Documents/abc')}')`,
      headers:
        { accept: 'application/json;odata=nometadata' },
      responseType: 'json'
    }));
  });

  it('should correctly pass params to request (sites/test1)', async () => {
    const request: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/test1',
        parentFolderUrl: 'Shared Documents',
        name: 'abc'
      }
    });
    assert(request.calledWith({
      url: `https://contoso.sharepoint.com/sites/test1/_api/web/folders/addUsingPath(decodedUrl='${formatting.encodeQueryParameter('/sites/test1/Shared Documents/abc')}')`,
      headers:
        { accept: 'application/json;odata=nometadata' },
      responseType: 'json'
    }));
  });

  it('creates a folder with a specific color by color number', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/foldercoloring/createfolder(DecodedUrl='${formatting.encodeQueryParameter(`${parentFolder}/${folderName}`)}', overwrite=false)`) {
        return addResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, parentFolderUrl: parentFolder, name: folderName, color: colorNumber } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      coloringInformation: {
        ColorHex: `${colorNumber}`
      }
    });
  });

  it('creates a folder with a specific color by color name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/foldercoloring/createfolder(DecodedUrl='${formatting.encodeQueryParameter(`${parentFolder}/${folderName}`)}', overwrite=false)`) {
        return addResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, parentFolderUrl: parentFolder, name: folderName, color: colorName } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      coloringInformation: {
        ColorHex: colorNumber
      }
    });
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', parentFolderUrl: parentFolder, name: folderName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and parentFolderUrl specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, parentFolderUrl: parentFolder, name: folderName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if color is passed as string and color is not a valid color', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, parentFolderUrl: parentFolder, name: folderName, color: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if color is passed as string and color is a valid color', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, parentFolderUrl: parentFolder, name: folderName, color: colorName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if color is passed as number and color is a valid color', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, parentFolderUrl: parentFolder, name: folderName, color: colorNumber } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
