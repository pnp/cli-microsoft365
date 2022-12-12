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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./web-remove');

describe(commands.WEB_REMOVE, () => {
  let log: any[];
  let requests: any[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => {});
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
    requests = [];
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: true };
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
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.WEB_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('should fail validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        url: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required options are specified', async () => {
    const actual = await command.validate({
      options: {
        url: "https://contoso.sharepoint.com/subsite"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should prompt before deleting subsite when confirmation argument not passed', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });
    
    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/subsite' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('deletes web successfully without prompting with confirmation argument', async () => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        url: "https://contoso.sharepoint.com/subsite",
        confirm: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web`) > -1 &&
        r.headers['X-HTTP-Method'] === 'DELETE' &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);

  });

  it('deletes web successfully when prompt confirmed', async () => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    
    await command.action(logger, {
      options: {
        url: "https://contoso.sharepoint.com/subsite"
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web`) > -1 &&
        r.headers['X-HTTP-Method'] === 'DELETE' &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('deletes web successfully without prompting with confirmation argument (verbose)', async () => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        verbose: true,
        url: "https://contoso.sharepoint.com/subsite",
        confirm: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web`) > -1 &&
        r.headers['X-HTTP-Method'] === 'DELETE' &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('deletes web successfully without prompting with confirmation argument (debug)', async () => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        url: "https://contoso.sharepoint.com/subsite",
        confirm: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web`) > -1 &&
        r.headers['X-HTTP-Method'] === 'DELETE' &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('handles error when deleting web', async () => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {
      url: "https://contoso.sharepoint.com/subsite",
      confirm: true } } as any), new CommandError('An error has occurred'));
  });
});
