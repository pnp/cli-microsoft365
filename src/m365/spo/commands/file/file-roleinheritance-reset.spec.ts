import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./file-roleinheritance-reset');
import * as SpoFileGetCommand from './file-get';

describe(commands.FILE_ROLEINHERITANCE_RESET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.prompt,
      request.get,
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.FILE_ROLEINHERITANCE_RESET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['fileId', 'fileUrl']]);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: 'b2307a39-e878-458b-bc90-03bc578531d6', confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: 'b2307a39-e878-458b-bc90-03bc578531d6', confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: 'foo', confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the fileId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: '0cd891ef-afce-4e55-b836-fce03286cccf', confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before resetting role inheritance for the file when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        fileId: 'b2307a39-e878-458b-bc90-03bc578531d6'
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts resetting role inheritance for the file when confirm option is not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        fileId: 'b2307a39-e878-458b-bc90-03bc578531d6'
      }
    });

    assert(postSpy.notCalled);
  });

  it('reset role inheritance on file by relative URL (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativeUrl('/sites/project-x/documents/Test1.docx')/ListItemAllFields/resetroleinheritance`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        fileUrl: '/sites/project-x/documents/Test1.docx',
        confirm: true
      }
    });
  });

  it('reset role inheritance on file by Id when prompt confirmed', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoFileGetCommand) {
        return Promise.resolve({
          stdout: '{"LinkingUri": "https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866","Name": "Test1.docx","ServerRelativeUrl": "/sites/project-x/documents/Test1.docx","UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativeUrl`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        fileId: 'b2307a39-e878-458b-bc90-03bc578531d6'
      }
    });
  });

  it('correctly handles error when resetting file role inheritance', async () => {
    const err = 'request rejected';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativeUrl('/sites/project-x/documents/Test1.docx')/ListItemAllFields/resetroleinheritance`) > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        fileUrl: '/sites/project-x/documents/Test1.docx',
        confirm: true
      }
    });
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
