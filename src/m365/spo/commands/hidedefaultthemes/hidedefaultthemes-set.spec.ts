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
const command: Command = require('./hidedefaultthemes-set');

describe(commands.HIDEDEFAULTTHEMES_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    requests = [];
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HIDEDEFAULTTHEMES_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets the value of the HideDefaultThemes setting', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('/_api/thememanager/SetHideDefaultThemes') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        hideDefaultThemes: true
      }
    });

    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/thememanager/SetHideDefaultThemes`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('sets the value of the HideDefaultThemes setting (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('/_api/thememanager/SetHideDefaultThemes') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        hideDefaultThemes: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/thememanager/SetHideDefaultThemes`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });

    assert(correctRequestIssued);
  });

  it('handles error when setting the value of the HideDefaultThemes setting', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('/_api/thememanager/SetHideDefaultThemes') > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        hideDefaultThemes: true
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if hideDefaultThemes is not set', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when hideDefaultThemes is true', async () => {
    const actual = await command.validate({ options: { hideDefaultThemes: true } }, commandInfo);
    assert(actual);
  });

  it('passes validation when hideDefaultThemes is false', async () => {
    const actual = await command.validate({ options: { hideDefaultThemes: false } }, commandInfo);
    assert(actual);
  });
});
