import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { cli } from '../../../cli/cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError } from '../../../Command.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import command, { options } from './cli-consent.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';

describe(commands.CONSENT, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: any;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(auth, 'restoreAuth').resolves();
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      cli.getTenant,
      cli.getClientId,
      (command as any).warn
    ]);
    auth.connection.active = false;
    auth.connection.appId = undefined;
    auth.connection.tenant = 'common';
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONSENT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows consent URL for VivaEngage permissions for a custom single-tenant app', async () => {
    sinon.stub(cli, 'getTenant').returns('fb5cb38f-ecdb-4c6a-a93b-b8cfd56b4a89');
    sinon.stub(cli, 'getClientId').returns('2587b55d-a41e-436d-bb1d-6223eb185dd4');
    await command.action(logger, { options: commandOptionsSchema.parse({ service: 'VivaEngage' }) });
    assert(loggerLogSpy.calledWith(`To consent permissions for executing VivaEngage commands, navigate in your web browser to https://login.microsoftonline.com/fb5cb38f-ecdb-4c6a-a93b-b8cfd56b4a89/oauth2/v2.0/authorize?client_id=2587b55d-a41e-436d-bb1d-6223eb185dd4&response_type=code&scope=https%3A%2F%2Fapi.yammer.com%2Fuser_impersonation`));
  });

  it('shows consent URL using the app ID and tenant of the active connection', async () => {
    sinon.stub(cli, 'getClientId').returns(undefined);
    auth.connection.active = true;
    auth.connection.appId = '31359c7f-bd7e-475c-86db-fdb8c937548e';
    auth.connection.tenant = 'fb5cb38f-ecdb-4c6a-a93b-b8cfd56b4a89';
    await command.action(logger, { options: commandOptionsSchema.parse({ service: 'VivaEngage' }) });
    assert(loggerLogSpy.calledWith(`To consent permissions for executing VivaEngage commands, navigate in your web browser to https://login.microsoftonline.com/fb5cb38f-ecdb-4c6a-a93b-b8cfd56b4a89/oauth2/v2.0/authorize?client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&response_type=code&scope=https%3A%2F%2Fapi.yammer.com%2Fuser_impersonation`));
  });

  it('throws an error when not logged in and no client ID is stored in the settings', async () => {
    sinon.stub(cli, 'getClientId').returns(undefined);
    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ service: 'VivaEngage' }) }),
      new CommandError(`Unable to determine the client ID of the Microsoft Entra application used by the CLI. Log in to Microsoft 365 first or store the client ID using 'm365 cli config set --key clientId --value <appId>'.`));
  });

  it('fails validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ service: 'VivaEngage', unknownOption: 'value' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if specified service is invalid ', () => {
    const actual = commandOptionsSchema.safeParse({ service: 'invalid' });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if service is set to VivaEngage ', () => {
    const actual = commandOptionsSchema.safeParse({ service: 'VivaEngage' });
    assert.strictEqual(actual.success, true);
  });
});
