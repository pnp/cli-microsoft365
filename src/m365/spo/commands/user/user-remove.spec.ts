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
import commands from '../../commands.js';
import command from './user-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.USER_REMOVE, () => {
  let cli: Cli;
  let log: any[];
  let requests: any[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

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
    requests = [];
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
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
    assert.strictEqual(command.name, commands.USER_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id or loginName options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id or loginname options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: 10,
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'foo',
        id: 10
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should prompt before removing user using id from web when confirmation argument not passed ', async () => {
    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/subsite',
        id: 10
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('should prompt before removing user using login name from web when confirmation argument not passed ', async () => {
    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/subsite',
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('removes user by id successfully without prompting with confirmation argument', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10,
        force: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`_api/web/siteusers/removebyid(10)`) > -1 &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user by login name successfully without prompting with confirmation argument', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if (opts.url === "https://contoso.sharepoint.com/subsite/_api/web/siteusers/removeByLoginName('i%3A0%23.f%7Cmembership%7Cparker%40tenant.onmicrosoft.com')") {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        loginName: "i:0#.f|membership|parker@tenant.onmicrosoft.com",
        force: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`_api/web/siteusers/removeByLoginName('i%3A0%23.f%7Cmembership%7Cparker%40tenant.onmicrosoft.com')`) > -1 &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user by id successfully from web when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        return true;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`_api/web/siteusers/removebyid(10)`) > -1 &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user by login name successfully from web when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`_api/web/siteusers/removeByLoginName`) > -1) {
        return true;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`_api/web/siteusers/removeByLoginName`) > -1 &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user from web successfully without prompting with confirmation argument (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10,
        force: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`_api/web/siteusers/removebyid(10)`) > -1 &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes user from web successfully without prompting with confirmation argument (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        return true;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10,
        force: true
      }
    });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`_api/web/siteusers/removebyid(10)`) > -1 &&
        r.headers['accept'] === 'application/json;odata=nometadata') {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('handles error when removing using from web', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        throw 'An error has occurred';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10,
        force: true
      }
    } as any), new CommandError('An error has occurred'));
  });
});
