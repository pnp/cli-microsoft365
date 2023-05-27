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
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
const command: Command = require('./file-checkout-undo');

describe(commands.FILE_CHECKOUT_UNDO, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/projects';
  const fileId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';
  const fileUrl = '/sites/projects/shared documents/test.docx';

  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    cli = Cli.getInstance();
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((_, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_CHECKOUT_UNDO);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('undos checkout for file retrieved by fileId when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/getFileById('${fileId}')/undocheckout`) {
        return;
      }

      throw 'Invalid request';
    });
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId, verbose: true } });
    assert(postStub.called);
  });

  it('undos checkout for file retrieved by fileUrl', async () => {
    const serverRelativePath = urlUtil.getServerRelativePath(webUrl, fileUrl);
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/getFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/undocheckout`) {
        return;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, confirm: true, verbose: true } });
    assert(postStub.called);
  });

  it('handles error when file is not checked out', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/getFileById('${fileId}')/undocheckout`) {
        throw {
          error: {
            'odata.error': {
              code: '-2147024738, Microsoft.SharePoint.SPFileCheckOutException',
              message: {
                lang: 'en-US',
                value: 'The file "Shared Documents/4.docx" is not checked out.'
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, fileId: fileId, confirm: true, verbose: true } }), new CommandError('The file "Shared Documents/4.docx" is not checked out.'));
  });

  it('prompts before undoing checkout when confirmation argument not passed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts undoing checkout when prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { webUrl: webUrl, id: fileId } });
    assert(postStub.notCalled);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', fileId: fileId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and fileId is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
