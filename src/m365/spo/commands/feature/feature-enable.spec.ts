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
import command from './feature-enable.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.FEATURE_ENABLE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    requests = [];
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FEATURE_ENABLE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('configures scope as string option', () => {
    const types = command.types;
    ['s', 'scope'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('enables web feature (scope not defined, so defaults to web), no force', async () => {
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/add(featureId=guid'b2307a39-e878-458b-bc90-03bc578531d6',force=false)`;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(requestUrl) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    try {
      await command.action(logger, { options: { debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(requestUrl) > -1 && r.headers.accept && r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
    finally {
      sinonUtil.restore(request.post);
    }
  });

  it('enables site feature, force', async () => {
    const requestUrl = `https://contoso.sharepoint.com/_api/site/features/add(featureId=guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3',force=true)`;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(requestUrl) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    try {
      await command.action(logger, { options: { debug: true, id: '915c240e-a6cc-49b8-8b2c-0bff8b553ed3', webUrl: 'https://contoso.sharepoint.com', scope: 'site', force: true } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(requestUrl) > -1 && r.headers.accept && r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      assert(correctRequestIssued);
    }
    finally {
      sinonUtil.restore(request.post);
    }
  });

  it('correctly handles error when enabling feature', async () => {
    const id = '780ac353-eaf8-4ac2-8c47-536d93c03fd6';
    const err = {
      error: {
        'odata.error': {
          message: {
            value: `Feature with Id '${id}' is not installed in this farm, and cannot be added to this scope.`
          }
        }
      }
    };

    sinon.stub(request, 'post').rejects(err);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: id,
        scope: 'web'
      }
    }), new CommandError(err.error['odata.error'].message.value));
  });

  it('supports specifying scope', () => {
    const options = command.options;
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[scope]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: 'invalid',
        scope: 'Web'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options:
      {
        webUrl: 'foo',
        id: '00bfea71-5932-4f9c-ad71-1557e5751100'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the required options specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        id: "00bfea71-5932-4f9c-ad71-1557e5751100"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Site', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        id: "00bfea71-5932-4f9c-ad71-1557e5751100",
        scope: 'Site'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Web', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        id: "00bfea71-5932-4f9c-ad71-1557e5751100",
        scope: 'Web'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid string scope', async () => {
    const scope = 'foo';
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        id: "00bfea71-5932-4f9c-ad71-1557e5751100",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', async () => {
    const actual = await command.validate(
      {
        options:
        {
          id: "00bfea71-5932-4f9c-ad71-1557e5751100",
          webUrl: "https://contoso.sharepoint.com"
        }
      }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
