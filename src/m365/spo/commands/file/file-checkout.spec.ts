import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./file-checkout');

describe(commands.FILE_CHECKOUT, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const stubPostResponses: any = (getFileByServerRelativeUrlResp: any = null, getFileByIdResp: any = null) => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if (getFileByServerRelativeUrlResp) {
        return getFileByServerRelativeUrlResp;
      }
      else {
        if ((opts.url as string).indexOf('/_api/web/GetFileByServerRelativeUrl(') > -1) {
          return Promise.resolve();
        }
      }

      if (getFileByIdResp) {
        return getFileByIdResp;
      }
      else {
        if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });
  };

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => { return defaultValue; }));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_CHECKOUT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('command correctly handles file get reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 'f09c4efe-b8c0-4e89-a166-03418661b89b'
      }
    }), new CommandError(err));
  });

  it('should handle checked out by someone else file', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575306, Microsoft.SharePoint.SPFileCheckOutException", "message": { "lang": "en-US", "value": "The file \"https://contoso.sharepoint.com/sites/xx/Shared Documents/abc.txt\" is checked out for editing by i:0#.f|membership|xx" } } });
    const getFileByServerRelativeUrlResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(getFileByServerRelativeUrlResp);

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }), new CommandError(expectedError));
  });

  it('should handle file does not exist', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File Not Found." } } });
    const getFileByIdResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, getFileByIdResp);

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }), new CommandError(expectedError));
  });

  it('should call the correct API url when UniqueId option is passed', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    await command.action(logger, {
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/checkout');
  });

  it('should call the correct API url when URL option is passed', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        url: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx')/checkout");
  });

  it('should call the correct API url when tenant root URL option is passed', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        url: '/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/_api/web/GetFileByServerRelativeUrl('%2FDocuments%2FTest1.docx')/checkout");
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

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the id or url option not specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and url options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', url: '/sites/project-x/documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
