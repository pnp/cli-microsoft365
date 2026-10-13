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
import command, { options } from './feature-enable.js';

describe(commands.FEATURE_ENABLE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()! as typeof options;
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
    assert.strictEqual(command.name, commands.FEATURE_ENABLE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('enables web feature (scope not defined, so defaults to web), no force', async () => {
    const requestUrl = `https://contoso.sharepoint.com/_api/Web/features/add(featureId=guid'b2307a39-e878-458b-bc90-03bc578531d6',force=false)`;
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
      await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' }) });
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
    const requestUrl = `https://contoso.sharepoint.com/_api/Site/features/add(featureId=guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3',force=true)`;
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
      await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: '915c240e-a6cc-49b8-8b2c-0bff8b553ed3', webUrl: 'https://contoso.sharepoint.com', scope: 'site', force: true }) });
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
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        id: id,
        scope: 'web'
      })
    }), new CommandError(err.error['odata.error'].message.value));
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      id: 'invalid',
      scope: 'Web'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, 'invalid is not a valid GUID for option id.');
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'foo',
      id: '00bfea71-5932-4f9c-ad71-1557e5751100'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('passes validation when the required options specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      id: "00bfea71-5932-4f9c-ad71-1557e5751100"
    });
    assert.strictEqual(actual.success, true);
  });

  it('accepts scope to be Site', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      id: "00bfea71-5932-4f9c-ad71-1557e5751100",
      scope: 'Site'
    });
    assert.strictEqual(actual.success, true);
  });

  it('accepts scope to be Web', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      id: "00bfea71-5932-4f9c-ad71-1557e5751100",
      scope: 'Web'
    });
    assert.strictEqual(actual.success, true);
  });

  it('rejects invalid string scope', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      id: "00bfea71-5932-4f9c-ad71-1557e5751100",
      scope: 'foo'
    });
    assert.strictEqual(actual.success, false);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: "00bfea71-5932-4f9c-ad71-1557e5751100",
      webUrl: "https://contoso.sharepoint.com"
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      id: "00bfea71-5932-4f9c-ad71-1557e5751100",
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
