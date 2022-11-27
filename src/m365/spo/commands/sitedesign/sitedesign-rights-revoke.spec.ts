import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./sitedesign-rights-revoke');

describe(commands.SITEDESIGN_RIGHTS_REVOKE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_RIGHTS_REVOKE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('revokes access to the specified site design from the specified principal without prompting for confirmation when confirm option specified', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
          principalNames: ['PattiF']
        })) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { confirm: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF' } });
    assert(loggerLogSpy.notCalled);
  });

  it('revokes access to the specified site design from the specified principals', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
          principalNames: ['PattiF', 'AdeleV']
        })) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { confirm: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF,AdeleV' } });
    assert(loggerLogSpy.notCalled);
  });

  it('revokes access to the specified site design from the principals specified via email', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
          principalNames: ['PattiF@contoso.com', 'AdeleV@contoso.com']
        })) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { confirm: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF@contoso.com,AdeleV@contoso.com' } });
    assert(loggerLogSpy.notCalled);
  });

  it('revokes access to the specified site design from the specified principals separated with an extra space', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
          principalNames: ['PattiF@contoso.com', 'AdeleV@contoso.com']
        })) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { confirm: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF@contoso.com, AdeleV@contoso.com' } });
    assert(loggerLogSpy.notCalled);
  });

  it('prompts before revoking access to the specified site design when confirm option not passed', async () => {
    await command.action(logger, { options: { siteDesignId: 'b2307a39-e878-458b-bc90-03bc578531d6', principals: 'PattiF' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts revoking access to the site design when prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    await command.action(logger, { options: { siteDesignId: 'b2307a39-e878-458b-bc90-03bc578531d6', principals: 'PattiF' } });
    assert(postSpy.notCalled);
  });

  it('revokes site design access when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(() => Promise.resolve());
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, { options: { siteDesignId: 'b2307a39-e878-458b-bc90-03bc578531d6', principals: 'PattiF' } });
    assert(postStub.called);
  });

  it('correctly handles error when site script not found', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('File Not Found.');
    });

    await assert.rejects(command.action(logger, { options: { confirm: true, siteDesignId: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', principals: 'PattiF' } } as any), new CommandError('File Not Found.'));
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
      if (o.option.indexOf('--confirm') > -1) {
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