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
import { urlUtil } from '../../../../utils/urlUtil.js';
import commands from '../../commands.js';
import command from './file-checkout-undo.js';

describe(commands.FILE_CHECKOUT_UNDO, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/projects';
  const fileId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';
  const fileUrl = '/sites/projects/shared documents/test.docx';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_CHECKOUT_UNDO);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('undoes checkout for file retrieved by fileId when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/getFileById('${fileId}')/undocheckout`) {
        return;
      }

      throw 'Invalid request';
    });
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId, verbose: true } });
    assert(postStub.called);
  });

  it('undoes checkout for file retrieved by fileUrl', async () => {
    const serverRelativePath = urlUtil.getServerRelativePath(webUrl, fileUrl);
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/getFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/undocheckout`) {
        return;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, force: true, verbose: true } });
    assert(postStub.called);
  });

  it('undoes checkout for file retrieved by site-relative url', async () => {
    const siteRelativeUrl = '/Shared Documents/Test.docx';
    const serverRelativePath = urlUtil.getServerRelativePath(webUrl, siteRelativeUrl);
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/getFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/undocheckout`) {
        return;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { webUrl: webUrl, fileUrl: siteRelativeUrl, force: true, verbose: true } });
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
    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, fileId: fileId, force: true, verbose: true } }), new CommandError('The file "Shared Documents/4.docx" is not checked out.'));
  });

  it('prompts before undoing checkout when confirmation argument not passed', async () => {
    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId } });

    assert(promptIssued);
  });

  it('aborts undoing checkout when prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
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
