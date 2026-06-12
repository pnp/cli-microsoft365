import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandError } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './connection-urltoitemresolver-add.js';

describe(commands.CONNECTION_URLTOITEMRESOLVER_ADD, () => {
  let logger: Logger;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    const commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
  });

  beforeEach(() => {
    logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONNECTION_URLTOITEMRESOLVER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds item to URL resolver to an existing external connection', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/conn`) {
        return {};
      }
      throw 'Invalid request';
    });
    const options: any = {
      externalConnectionId: 'conn',
      baseUrls: 'https://contoso.com',
      urlPattern: '/(?<id>.*)',
      itemId: '{id}',
      priority: 1
    };
    await command.action(logger, { options: commandOptionsSchema.parse(options) });
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'patch').callsFake(() => {
      throw {
        "error": {
          "code": "Error",
          "message": "An error has occurred",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      };
    });

    const options: any = {
      externalConnectionId: 'conn',
      baseUrls: 'https://contoso.com',
      urlPattern: '/(?<id>.*)',
      itemId: '{id}',
      priority: 1
    };

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse(options) }),
      new CommandError(`An error has occurred`));
  });

  it('passes validation with all required options', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: 'conn',
      baseUrls: 'https://contoso.com',
      urlPattern: '/(?<id>.*)',
      itemId: '{id}',
      priority: 1
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      externalConnectionId: 'conn',
      baseUrls: 'https://contoso.com',
      urlPattern: '/(?<id>.*)',
      itemId: '{id}',
      priority: 1,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
