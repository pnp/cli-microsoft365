import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './sitescript-remove.js';
import { CentralizedTestSetup, initializeTestSetup } from '../../../../utils/tests.js';

describe(commands.SITESCRIPT_REMOVE, () => {
  let commandInfo: CommandInfo;
  let promptOptions: any;
  let testSetup: CentralizedTestSetup;

  before(() => {
    testSetup = initializeTestSetup();
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    testSetup.runBeforeEachHookDefaults();
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    testSetup.runAfterEachHookDefaults();
    sinonUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    testSetup.runAfterHookDefaults();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITESCRIPT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the specified site script without prompting for confirmation when confirm option specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(testSetup.logger, { options: { force: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } });
  });

  it('prompts before removing the specified site script when confirm option not passed', async () => {
    await command.action(testSetup.logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('aborts removing site script when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(testSetup.logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
    assert(postSpy.notCalled);
  });

  it('removes the app when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(testSetup.logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6' } });
    assert(postStub.called);
  });

  it('correctly handles error when site script not found', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });

    await assert.rejects(command.action(testSetup.logger, { options: { force: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } } as any), new CommandError('File Not Found.'));
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
