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
import command, { options } from './roledefinition-add.js';

describe(commands.ROLEDEFINITION_ADD, () => {
  let log: any[];
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
    assert.strictEqual(command.name, commands.ROLEDEFINITION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo', name: 'abc' });
    assert.strictEqual(actual.success, false);
    assert.ok(actual.error.issues[0].message.includes(`'foo' is not a valid SharePoint Online site URL.`));
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: 'abc' });
    assert.strictEqual(actual.success, true);
  });

  it('fails if non existing PermissionKind rights specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: 'abc', rights: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('has correct PermissionKind rights specified', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: 'abc', rights: 'FullMask' });
    assert.strictEqual(actual.success, true);
  });

  it('adds role definition to web with name, description and right', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/roledefinitions') {
        return '';
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        name: 'test',
        description: 'test',
        rights: 'FullMask'
      })
    });
  });

  it('adds role definition to web with name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/web/roledefinitions') {
        return '';
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        name: 'test'
      })
    });
  });

  it('handles reject request correctly', async () => {
    const err = 'request rejected';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/roledefinitions') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        name: 'test'
      })
    }), new CommandError(err));
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', name: 'test', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('returns a descriptive validation error for an invalid PermissionKind right', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      name: 'abc',
      rights: 'FullMask, invalidRight'
    });

    assert.strictEqual(actual.success, false);

    if (actual.success) {
      assert.fail('Expected validation to fail');
    }

    assert.ok(actual.error.issues[0].message.includes(
      "Rights option ' invalidRight' is not recognized as valid PermissionKind choice."
    ));
    assert.ok(actual.error.issues[0].message.includes(
      'Please note it is case-sensitive. Allowed values are '
    ));
    assert.ok(actual.error.issues[0].message.includes('FullMask'));
  });
});
