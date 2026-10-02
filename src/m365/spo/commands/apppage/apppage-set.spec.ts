import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { telemetry } from '../../../../telemetry.js';
import commands from '../../commands.js';
import command, { options } from './apppage-set.js';

describe(commands.APPPAGE_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    assert.strictEqual(command.name, commands.APPPAGE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails to update the single-part app page if request is rejected', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url as string).indexOf('_api/sitepages/Pages/UpdateFullPageApp') > -1 &&
        opts.data.serverRelativeUrl.indexOf('failme')) {
        throw 'Failed to update the single-part app page';
      }
      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger,
      {
        options: commandOptionsSchema.parse({
          name: 'failme',
          webUrl: 'https://contoso.sharepoint.com/',
          webPartData: JSON.stringify({})
        })
      }), new CommandError('Failed to update the single-part app page'));
  });

  it('updates the single-part app page', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url as string).indexOf('_api/sitepages/Pages/UpdateFullPageApp') > -1) {
        return;
      }
      throw 'Invalid request';
    });
    await command.action(logger,
      {
        options: commandOptionsSchema.parse({
          name: 'demo',
          webUrl: 'https://contoso.sharepoint.com/teams/sales',
          webPartData: JSON.stringify({})
        })
      });
  });

  it('fails validation if name not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webPartData: JSON.stringify({ abc: 'def' }),
      webUrl: 'https://contoso.sharepoint.com'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if webPartData not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Contoso.aspx',
      webUrl: 'https://contoso.sharepoint.com'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if webUrl not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webPartData: JSON.stringify({ abc: 'def' }),
      name: 'page.aspx'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if webPartData is not a valid JSON string', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Contoso.aspx',
      webUrl: 'https://contoso',
      webPartData: 'abc'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Contoso.aspx',
      webPartData: '{}',
      webUrl: 'https://contoso.sharepoint.com',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('validation passes on all required options', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Contoso.aspx',
      webPartData: '{}',
      webUrl: 'https://contoso.sharepoint.com'
    });
    assert.strictEqual(actual.success, true);
  });
});
