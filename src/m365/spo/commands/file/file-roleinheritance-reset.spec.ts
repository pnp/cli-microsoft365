import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./file-roleinheritance-reset');
import * as SpoFileGetCommand from './file-get';
import { session } from '../../../../utils/session';
import { pid } from '../../../../utils/pid';

describe(commands.FILE_ROLEINHERITANCE_RESET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const fileUrl = '/sites/project-x/documents/Test1.docx';
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_ROLEINHERITANCE_RESET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'foo', force: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and fileId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: '0cd891ef-afce-4e55-b836-fce03286cccf', force: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before resetting role inheritance for the file when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId
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
        webUrl: webUrl,
        fileId: fileId
      }
    });

    assert(postSpy.notCalled);
  });

  it('resets role inheritance on file by relative URL (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/resetroleinheritance`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        force: true
      }
    });
  });

  it('resets role inheritance on file by Id when prompt confirmed', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoFileGetCommand) {
        return ({
          stdout: '{"LinkingUri": "https://contoso.sharepoint.com/sites/project-x/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866","Name": "Test1.docx","ServerRelativeUrl": "/sites/project-x/documents/Test1.docx","UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"}'
        });
      }

      throw new CommandError('Unknown case');
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/resetroleinheritance`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        fileId: fileId
      }
    });
  });

  it('correctly handles error when resetting file role inheritance', async () => {
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
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        force: true
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});
