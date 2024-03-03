import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
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
import command from './sitedesign-rights-revoke.js';

describe(commands.SITEDESIGN_RIGHTS_REVOKE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITEDESIGN_RIGHTS_REVOKE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('revokes access to the specified site design from the specified principal without prompting for confirmation when force option specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
          principalNames: ['PattiF']
        })) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF' } });
    assert(loggerLogSpy.notCalled);
  });

  it('revokes access to the specified site design from the specified principals', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
          principalNames: ['PattiF', 'AdeleV']
        })) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF,AdeleV' } });
    assert(loggerLogSpy.notCalled);
  });

  it('revokes access to the specified site design from the principals specified via email', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
          principalNames: ['PattiF@contoso.com', 'AdeleV@contoso.com']
        })) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF@contoso.com,AdeleV@contoso.com' } });
    assert(loggerLogSpy.notCalled);
  });

  it('revokes access to the specified site design from the specified principals separated with an extra space', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
          principalNames: ['PattiF@contoso.com', 'AdeleV@contoso.com']
        })) {
        return {
          "odata.null": true
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF@contoso.com, AdeleV@contoso.com' } });
    assert(loggerLogSpy.notCalled);
  });

  it('prompts before revoking access to the specified site design when force option not passed', async () => {
    await command.action(logger, { options: { siteDesignId: 'b2307a39-e878-458b-bc90-03bc578531d6', principals: 'PattiF' } });

    assert(promptIssued);
  });

  it('aborts revoking access to the site design when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    await command.action(logger, { options: { siteDesignId: 'b2307a39-e878-458b-bc90-03bc578531d6', principals: 'PattiF' } });
    assert(postSpy.notCalled);
  });

  it('revokes site design access when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { siteDesignId: 'b2307a39-e878-458b-bc90-03bc578531d6', principals: 'PattiF' } });
    assert(postStub.called);
  });

  it('correctly handles error when site script not found', async () => {
    sinon.stub(request, 'post').rejects(new Error('File Not Found.'));

    await assert.rejects(command.action(logger, { options: { force: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF' } } as any), new CommandError('File Not Found.'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteDesignId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying principals', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--principals') > -1) {
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
    const actual = await command.validate({ options: { siteDesignId: 'abc', principals: 'PattiF' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the all required parameters are specified', async () => {
    const actual = await command.validate({ options: { siteDesignId: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a', principals: 'PattiF' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
