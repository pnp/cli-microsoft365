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
import commands from '../../commands.js';
import command from './feature-disable.js';

describe(commands.FEATURE_DISABLE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FEATURE_DISABLE);
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

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: 'invalid',
        scope: 'Site'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is not site|web', async () => {
    const scope = 'list';
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        id: "780ac353-eaf8-4ac2-8c47-536d93c03fd6",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid Feature scope. Allowed values are: Site, Web.`);
  });

  it('passes validation if webUrl and id is correct', async () => {
    const actual = await command.validate({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        id: "780ac353-eaf8-4ac2-8c47-536d93c03fd6"
      }
    }, commandInfo);

    assert.strictEqual(actual, true);
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

  it('disables web feature (scope not defined, so defaults to web), no force', async () => {
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/remove(featureId=guid'780ac353-eaf8-4ac2-8c47-536d93c03fd6',force=false)`;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      {
        requests.push(opts);

        if ((opts.url as string).indexOf(requestUrl) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      }
    });

    try {
      await command.action(logger, { options: { debug: true, id: '780ac353-eaf8-4ac2-8c47-536d93c03fd6', webUrl: 'https://contoso.sharepoint.com' } });
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

  it('disables web feature (scope not defined, so defaults to web), with force', async () => {
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/remove(featureId=guid'780ac353-eaf8-4ac2-8c47-536d93c03fd6',force=true)`;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      {
        requests.push(opts);

        if ((opts.url as string).indexOf(requestUrl) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      }
    });

    try {
      await command.action(logger, { options: { debug: true, id: '780ac353-eaf8-4ac2-8c47-536d93c03fd6', webUrl: 'https://contoso.sharepoint.com', force: true } });
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

  it('disables site feature (scope explicitly set), no force', async () => {
    const requestUrl = `https://contoso.sharepoint.com/_api/site/features/remove(featureId=guid'780ac353-eaf8-4ac2-8c47-536d93c03fd6',force=false)`;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      {
        requests.push(opts);

        if ((opts.url as string).indexOf(requestUrl) > -1) {
          if (opts.headers &&
            opts.headers.accept &&
            (opts.headers.accept as string).indexOf('application/json') === 0) {
            return;
          }
        }

        throw 'Invalid request';
      }
    });

    try {
      await command.action(logger, { options: { debug: true, id: '780ac353-eaf8-4ac2-8c47-536d93c03fd6', webUrl: 'https://contoso.sharepoint.com', scope: 'site' } });
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

  it('correctly handles disable feature reject request', async () => {
    const id = '780ac353-eaf8-4ac2-8c47-536d93c03fd6';
    const err = {
      error: {
        'odata.error': {
          message: {
            value: `Feature '${id}' is not activated at this scope.`
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
});
