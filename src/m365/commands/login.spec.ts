import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { z } from 'zod';
import auth, { AuthType, CloudType } from '../../Auth.js';
import { CommandArgs, CommandError } from '../../Command.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import { cli } from '../../cli/cli.js';
import { telemetry } from '../../telemetry.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import commands from './commands.js';
import command from './login.js';

describe(commands.LOGIN, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '123',
      accessToken: 'abc'
    };
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
    sinon.stub(auth.connection, 'deactivate').callsFake(() => { });
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => {
      auth.connection.name = '028de82d-7fd9-476e-a9fd-be9714280ff3';
      auth.connection.identityName = 'alexw@contoso.com';
      auth.connection.identityId = '028de82d-7fd9-476e-a9fd-be9714280ff3';
      auth.connection.identityTenantId = 'db308122-52f3-4241-af92-1734aa6e2e50';
      return Promise.resolve('');
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      auth.connection.deactivate,
      auth.ensureAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LOGIN), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('in telemetry tracks the specified cloud', () => {
    const args: CommandArgs = { options: { cloud: 'USGov' } };
    const telemetryProperties = (command as any).getTelemetryProperties(args);
    assert.strictEqual(telemetryProperties.cloud, 'USGov');
  });

  it('logs in to Microsoft 365', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert(auth.connection.active);
  });

  it('logs in to Microsoft 365 (debug)', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) });
    assert(auth.connection.active);
  });

  it('logs in to Microsoft 365 using username and password when authType password set', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ authType: 'password', userName: 'user', password: 'password' }) });
    assert.strictEqual(auth.connection.authType, AuthType.Password, 'Incorrect authType set');
    assert.strictEqual(auth.connection.userName, 'user', 'Incorrect user name set');
    assert.strictEqual(auth.connection.password, 'password', 'Incorrect password set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateFile is provided', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    await command.action(logger, { options: commandOptionsSchema.parse({ authType: 'certificate', certificateFile: 'certificate' }) });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set with thumbprint', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    await command.action(logger, { options: commandOptionsSchema.parse({ authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' }) });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateBase64Encoded is provided', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    await command.action(logger, { options: commandOptionsSchema.parse({ authType: 'certificate', certificateBase64Encoded: 'certificate', thumbprint: 'thumbprint' }) });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using system managed identity when authType identity set', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ authType: 'identity', userName: 'ac9fbed5-804c-4362-a369-21a4ec51109e' }) });
    assert.strictEqual(auth.connection.authType, AuthType.Identity, 'Incorrect authType set');
    assert.strictEqual(auth.connection.userName, 'ac9fbed5-804c-4362-a369-21a4ec51109e', 'Incorrect userName set');
  });

  it('logs in to Microsoft 365 using user-assigned managed identity when authType identity set', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ authType: 'identity' }) });
    assert.strictEqual(auth.connection.authType, AuthType.Identity, 'Incorrect authType set');
    assert.strictEqual(auth.connection.userName, undefined, 'Incorrect userName set');
  });
  it('logs in to Microsoft 365 using client secret authType "secret" set', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ authType: 'secret', secret: 'unBrEakaBle@123' }) });
    assert.strictEqual(auth.connection.authType, AuthType.Secret, 'Incorrect authType set');
    assert.strictEqual(auth.connection.secret, 'unBrEakaBle@123', 'Incorrect secret set');
  });

  it('logs in to Microsoft 365 using the specified cloud', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ cloud: 'USGov' }) });
    assert.strictEqual(auth.connection.cloudType, CloudType.USGov);
  });

  it('fails validation if invalid authType specified', () => {
    const actual = commandOptionsSchema.safeParse({ authType: 'invalid authType' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to password and userName and password not specified', () => {
    const actual = commandOptionsSchema.safeParse({ authType: 'password' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to password and userName not specified', () => {
    const actual = commandOptionsSchema.safeParse({ authType: 'password', password: 'password' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to password and password not specified', () => {
    const actual = commandOptionsSchema.safeParse({ authType: 'password', userName: 'user' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to certificate and both certificateFile and certificateBase64Encoded are specified', () => {
    const actual = commandOptionsSchema.safeParse({ authType: 'certificate', certificateFile: 'certificate', certificateBase64Encoded: 'certificateB64', thumbprint: 'thumbprint' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to certificate and neither certificateFile nor certificateBase64Encoded are specified', () => {
    const actual = commandOptionsSchema.safeParse({ authType: 'certificate' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to certificate and certificateFile does not exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = commandOptionsSchema.safeParse({ authType: 'certificate', certificateFile: 'certificate' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation cloud is set to an invalid value', () => {
    const actual = commandOptionsSchema.safeParse({ cloud: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if authType is set to certificate and certificateFile and thumbprint are specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = commandOptionsSchema.safeParse({ authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if authType is set to certificate and certificateFile are specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = commandOptionsSchema.safeParse({ authType: 'certificate', certificateFile: 'certificate' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if authType is set to password and userName and password specified', () => {
    const actual = commandOptionsSchema.safeParse({ authType: 'password', userName: 'user', password: 'password' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if authType is set to deviceCode and userName and password not specified', () => {
    const actual = commandOptionsSchema.safeParse({ authType: 'deviceCode' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if authType is not set and userName and password not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to Public', () => {
    const actual = commandOptionsSchema.safeParse({ cloud: 'Public' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to USGov', () => {
    const actual = commandOptionsSchema.safeParse({ cloud: 'USGov' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to USGovHigh', () => {
    const actual = commandOptionsSchema.safeParse({ cloud: 'USGovHigh' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to USGovDoD', () => {
    const actual = commandOptionsSchema.safeParse({ cloud: 'USGovDoD' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to China', () => {
    const actual = commandOptionsSchema.safeParse({ cloud: 'China' });
    assert.strictEqual(actual.success, true);
  });

  it('correctly handles error in device code auth flow', async () => {
    sinonUtil.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('Error'));
  });

  it('correctly handles error in device code auth flow (debug)', async () => {
    sinonUtil.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) }), new CommandError('Error'));
  });

  it('logs in to Microsoft 365 using browser authentication', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ authType: 'browser' }) });
    assert.strictEqual(auth.connection.authType, AuthType.Browser, 'Incorrect authType set');
  });

  it('correctly handles error when clearing persisted auth information', async () => {
    sinonUtil.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));

    try {
      await command.action(logger, { options: commandOptionsSchema.parse({}) });
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });

  it('correctly handles error when clearing persisted auth information (debug)', async () => {
    sinonUtil.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));

    try {
      await command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) });
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });

  it('correctly handles error when restoring auth information', async () => {
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    try {
      await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) } as any), new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });

  it('fails validation if authType is set to secret and secret option is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ authType: 'secret' });
    assert.strictEqual(actual.success, false);
  });
});
