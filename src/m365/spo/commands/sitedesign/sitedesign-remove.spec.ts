import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './sitedesign-remove.js';

describe(commands.SITEDESIGN_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the specified site design without prompting for confirmation when force option specified', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteDesign`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { force: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } });
    assert(loggerLogSpy.notCalled);
  });

  it('prompts before removing the specified site design when force option not passed', async () => {
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });

    assert(promptIssued);
  });

  it('aborts removing site design when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
    assert(postSpy.notCalled);
  });

  it('removes the app when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(() => Promise.resolve());

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
    assert(postStub.called);
  });

  it('correctly handles error when site design not found', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });
    });

    await assert.rejects(command.action(logger, { options: { force: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } } as any), new CommandError('File Not Found.'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirmation flag', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--force') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
