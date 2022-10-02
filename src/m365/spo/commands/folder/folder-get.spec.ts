import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./folder-get');

describe(commands.FOLDER_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let stubGetResponses: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;

    stubGetResponses = (getResp: any = null) => {
      return sinon.stub(request, 'get').callsFake((opts) => {
        if ((opts.url as string).indexOf('GetFolderByServerRelativeUrl') > -1 || (opts.url as string).indexOf('GetFolderById') > -1) {
          if (getResp) {
            return getResp;
          }
          else {
            return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/test1", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" });
          }
        }

        return Promise.reject('Invalid request');
      });
    };
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FOLDER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['folderUrl', 'id']]);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', folderUrl: '/Shared Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and folderUrl specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('should correctly handle folder get reject request', async () => {
    stubGetResponses(new Promise((resolve, reject) => { reject('error1'); }));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents'
      }
    } as any), new CommandError('error1'));
  });

  it('should show tip when folder get rejects with error code 500', async () => {
    sinon.stub(request, 'get').rejects({ statusCode: 500 });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents'
      }
    } as any), new CommandError('Please check the folder URL. Folder might not exist on the specified URL'));
  });

  it('should correctly handle folder get success request', async () => {
    stubGetResponses();

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents'
      }
    });
    assert(loggerLogSpy.lastCall.calledWith({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/test1", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" }));
  });

  it('should pass the correct id params to request', async () => {
    const request = stubGetResponses();

    await command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com',
        id: 'b2307a39-e878-458b-bc90-03bc578531d6'
      }
    });
    const lastCall: any = request.lastCall.args[0];
    assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderById(\'b2307a39-e878-458b-bc90-03bc578531d6\')');
  });

  it('should pass the correct url params to request', async () => {
    const request = stubGetResponses();

    await command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents'
      }
    });
    const lastCall: any = request.lastCall.args[0];
    assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl(\'%2FShared%20Documents\')');
  });

  it('should pass the correct url params to request (sites/test1)', async () => {
    const request = stubGetResponses();

    await command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com/sites/test1',
        folderUrl: 'Shared Documents/'
      }
    });
    const lastCall: any = request.lastCall.args[0];
    assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/sites/test1/_api/web/GetFolderByServerRelativeUrl(\'%2Fsites%2Ftest1%2FShared%20Documents\')');
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});