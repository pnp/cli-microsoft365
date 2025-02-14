import assert from 'assert';
import Configstore from 'configstore';
import fs from 'fs';
import sinon from 'sinon';
import { z } from 'zod';
import auth, { AuthType, CloudType } from '../../Auth.js';
import { CommandArgs, CommandError } from '../../Command.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import { cli } from '../../cli/cli.js';
import { settingsNames } from '../../settingsNames.js';
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
  let config: Configstore;
  let deactivateStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '123',
      accessToken: 'abc'
    };
    config = cli.getConfig();
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
    deactivateStub = sinon.stub(auth.connection, 'deactivate').callsFake(() => { });
    sinon.stub(auth, 'ensureAccessToken').callsFake(async () => {
      auth.connection.name = '028de82d-7fd9-476e-a9fd-be9714280ff3';
      auth.connection.identityName = 'alexw@contoso.com';
      auth.connection.identityId = '028de82d-7fd9-476e-a9fd-be9714280ff3';
      auth.connection.identityTenantId = 'db308122-52f3-4241-af92-1734aa6e2e50';
      return '';
    });
    sinon.stub(config, 'get').returns(undefined);
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      auth.connection.deactivate,
      auth.ensureAccessToken,
      config.get
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
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000'
      })
    });
    assert(auth.connection.active);
  });

  it('logs in to Microsoft 365 using appId and tenant set in CLI config', async () => {
    sinonUtil.restore(config.get);
    sinon.stub(config, 'get').callsFake(setting => {
      if (setting === settingsNames.clientId) {
        return '00000000-0000-0000-0000-000000000000';
      }
      else if (setting === settingsNames.tenantId) {
        return '00000000-0000-0000-0000-000000000000';
      }
      else {
        return undefined;
      }
    });
    await command.action(logger, {
      options: commandOptionsSchema.parse({})
    });
    assert(auth.connection.active);
  });

  it('logs in to Microsoft 365 (debug)', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        debug: true
      })
    });
    assert(auth.connection.active);
  });

  it('logs in to Microsoft 365 using username and password when authType password set', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'password',
        userName: 'user',
        password: 'password'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Password, 'Incorrect authType set');
    assert.strictEqual(auth.connection.userName, 'user', 'Incorrect user name set');
    assert.strictEqual(auth.connection.password, 'password', 'Incorrect password set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateFile is provided', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'certificate',
        certificateFile: 'certificate'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificate password is provided', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'certificate',
        certificateFile: 'certificate',
        password: 'p@$$w0rd'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.password, 'p@$$w0rd', 'Incorrect password set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificate password is empty', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'certificate',
        certificateFile: 'certificate',
        password: ''
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.password, '', 'Incorrect password set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set with thumbprint', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'certificate',
        certificateFile: 'certificate',
        thumbprint: 'thumbprint'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateBase64Encoded is provided', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'certificate',
        certificateBase64Encoded: 'certificate',
        thumbprint: 'thumbprint'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateBase64Encoded is set in CLI config', async () => {
    sinonUtil.restore(config.get);
    sinon.stub(config, 'get').callsFake(setting => {
      if (setting === settingsNames.clientCertificateBase64Encoded) {
        return 'certificate';
      }
      else {
        return undefined;
      }
    });
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'certificate',
        thumbprint: 'thumbprint'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and clientCertificatePassword is set in CLI config', async () => {
    sinonUtil.restore(config.get);
    sinon.stub(config, 'get').callsFake(setting => {
      if (setting === settingsNames.clientCertificateBase64Encoded) {
        return 'certificate';
      }
      if (setting === settingsNames.clientCertificatePassword) {
        return 'p@$$w0rd';
      }
      return undefined;
    });
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'certificate'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.password, 'p@$$w0rd', 'Incorrect password set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateFile is set in CLI config', async () => {
    sinonUtil.restore(config.get);
    sinon.stub(config, 'get').callsFake(setting => {
      if (setting === settingsNames.clientCertificateFile) {
        return 'certificate';
      }
      else {
        return undefined;
      }
    });
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'certificate',
        thumbprint: 'thumbprint'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using system managed identity when authType identity set', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'identity',
        userName: 'ac9fbed5-804c-4362-a369-21a4ec51109e'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Identity, 'Incorrect authType set');
    assert.strictEqual(auth.connection.userName, 'ac9fbed5-804c-4362-a369-21a4ec51109e', 'Incorrect userName set');
  });

  it('logs in to Microsoft 365 using user-assigned managed identity when authType identity set', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'identity'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Identity, 'Incorrect authType set');
    assert.strictEqual(auth.connection.userName, undefined, 'Incorrect userName set');
  });

  it('logs in to Microsoft 365 using system-assigned managed identity when authType identity set', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        authType: 'identity'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Identity, 'Incorrect authType set');
  });

  it('logs in to Microsoft 365 using client secret authType "secret" set', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'secret',
        secret: 'unBrEakaBle@123'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Secret, 'Incorrect authType set');
    assert.strictEqual(auth.connection.secret, 'unBrEakaBle@123', 'Incorrect secret set');
  });

  it('logs in to Microsoft 365 using client secret authType "secret" with secret set in CLI config', async () => {
    sinonUtil.restore(config.get);
    sinon.stub(config, 'get').callsFake(setting => {
      if (setting === settingsNames.clientSecret) {
        return 'unBrEakaBle@123';
      }
      else {
        return undefined;
      }
    });
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'secret'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Secret, 'Incorrect authType set');
    assert.strictEqual(auth.connection.secret, 'unBrEakaBle@123', 'Incorrect secret set');
  });

  it('logs in to Microsoft 365 using the specified cloud', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        cloud: 'USGov'
      })
    });
    assert.strictEqual(auth.connection.cloudType, CloudType.USGov);
  });

  it('fails validation if invalid authType specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'invalid authType'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to password and userName and password not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'password'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to password and userName not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'password',
      password: 'password'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to password and password not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'password',
      userName: 'user'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to certificate and both certificateFile and certificateBase64Encoded are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'certificate',
      certificateFile: 'certificate',
      certificateBase64Encoded: 'certificateB64',
      thumbprint: 'thumbprint'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to certificate and neither certificateFile nor certificateBase64Encoded are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'certificate'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if authType is set to certificate and certificateFile does not exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = commandOptionsSchema.safeParse({
      authType: 'certificate',
      certificateFile: 'certificate'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation cloud is set to an invalid value', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      cloud: 'invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if authType is set to certificate and certificateFile and thumbprint are specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'certificate',
      certificateFile: 'certificate',
      thumbprint: 'thumbprint'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if authType is set to certificate and certificateFile are specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'certificate',
      certificateFile: 'certificate'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if authType is set to password and userName and password specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'password',
      userName: 'user',
      password: 'password'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if authType is set to deviceCode and userName and password not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'deviceCode'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if authType is not set and userName and password not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to Public', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      cloud: 'Public'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to USGov', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      cloud: 'USGov'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to USGovHigh', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      cloud: 'USGovHigh'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to USGovDoD', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      cloud: 'USGovDoD'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when cloud is set to China', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      cloud: 'China'
    });
    assert.strictEqual(actual.success, true);
  });

  it('correctly handles error in device code auth flow', async () => {
    sinonUtil.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000'
      })
    }), new CommandError('Error'));
  });

  it('correctly handles error in device code auth flow (debug)', async () => {
    sinonUtil.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        debug: true
      })
    }), new CommandError('Error'));
  });

  it('logs in to Microsoft 365 using browser authentication', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        authType: 'browser'
      })
    });
    assert.strictEqual(auth.connection.authType, AuthType.Browser, 'Incorrect authType set');
  });

  it(`doesn't start the login flow when the CLI is signed in`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.DeviceCode,
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true
      })
    });

    assert(deactivateStub.notCalled);
  });

  it(`doesn't start the login flow if the CLI is signed in as a user`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.Password,
      userName: 'john.doe@contoso.com',
      password: 'password',
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true,
        authType: 'password',
        userName: 'john.doe@contoso.com',
        password: 'password'
      })
    });

    assert(deactivateStub.notCalled);
  });

  it(`doesn't start the login flow if the CLI is signed in using a certificate`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.Certificate,
      certificate: 'certificate',
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('certificate');

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true,
        authType: 'certificate',
        certificateFile: 'certificate'
      })
    });

    assert(deactivateStub.notCalled);
  });

  it(`doesn't start the login flow if the CLI is signed in using the specified app and to the specified tenant`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.DeviceCode,
      appId: '1cf21ca6-c8f0-4a21-839d-68a09d3a0f55',
      tenant: '973fce64-6409-4843-9328-c2cef0427f4e'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '1cf21ca6-c8f0-4a21-839d-68a09d3a0f55',
        tenant: '973fce64-6409-4843-9328-c2cef0427f4e',
        ensure: true,
        authType: 'deviceCode'
      })
    });

    assert(deactivateStub.notCalled);
  });

  it(`starts the login flow again when using a different auth type`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.Password,
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      userName: 'user@contoso.com',
      password: 'pass@word1'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true,
        authType: 'identity',
        userName: 'ac9fbed5-804c-4362-a369-21a4ec51109e'
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow again when using a different cloud type`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.DeviceCode,
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      cloudType: CloudType.Public
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true,
        cloud: 'USGov'
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow again when using a different app id`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      cloudType: CloudType.Public,
      appId: '1cf21ca6-c8f0-4a21-839d-68a09d3a0f55',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: 'b059efda-fc9d-49ec-b585-283f5b26202e',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow again when using a different tenant id`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      appId: '1cf21ca6-c8f0-4a21-839d-68a09d3a0f55',
      tenant: '973fce64-6409-4843-9328-c2cef0427f4e'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '1cf21ca6-c8f0-4a21-839d-68a09d3a0f55',
        tenant: '7f7993c9-ae48-413a-ae6b-d816a669f602',
        ensure: true
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow again when using a different username and authType password`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.Password,
      userName: 'user@contoso.com',
      password: 'password',
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true,
        authType: 'password',
        userName: 'user1@contoso.com',
        password: 'password'
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow again when using a different certificate file`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.Certificate,
      certificate: 'certificate',
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('certificate1');

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true,
        authType: 'certificate',
        certificateFile: 'certificate'
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow again when using a different username and authType identity`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.Identity,
      userName: 'ac9fbed5-804c-4362-a369-21a4ec51109e',
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true,
        authType: 'identity',
        userName: '1cf21ca6-c8f0-4a21-839d-68a09d3a0f55'
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow again when using a different secret`, async () => {
    const future = new Date();
    future.setSeconds(future.getSeconds() + 10);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.Secret,
      secret: 'topSeCr3t@007',
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: future.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true,
        authType: 'secret',
        secret: 'topSeCr3t@008'
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow when the access token expired (Date)`, async () => {
    const past = new Date();
    past.setSeconds(past.getSeconds() - 1);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.DeviceCode,
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: past,
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow when the access token expired (string)`, async () => {
    const past = new Date();
    past.setSeconds(past.getSeconds() - 1);
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.DeviceCode,
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: past.toISOString(),
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true
      })
    });

    assert(deactivateStub.called);
  });

  it(`starts the login flow when the access token has no expiration date (null)`, async () => {
    Object.assign(auth.connection, {
      active: true,
      authType: AuthType.DeviceCode,
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000'
    });
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: null,
      accessToken: 'abc'
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        appId: '00000000-0000-0000-0000-000000000000',
        tenant: '00000000-0000-0000-0000-000000000000',
        ensure: true
      })
    });

    assert(deactivateStub.called);
  });

  it('correctly handles error when clearing persisted auth information', async () => {
    sinonUtil.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));

    try {
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          appId: '00000000-0000-0000-0000-000000000000',
          tenant: '00000000-0000-0000-0000-000000000000'
        })
      });
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
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          appId: '00000000-0000-0000-0000-000000000000',
          tenant: '00000000-0000-0000-0000-000000000000',
          debug: true
        })
      });
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
      await assert.rejects(command.action(logger, {
        options: commandOptionsSchema.parse({
          appId: '00000000-0000-0000-0000-000000000000',
          tenant: '00000000-0000-0000-0000-000000000000',
          debug: true
        })
      } as any), new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });

  it('fails validation if authType is set to secret and secret option is not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      appId: '00000000-0000-0000-0000-000000000000',
      tenant: '00000000-0000-0000-0000-000000000000',
      authType: 'secret'
    });
    assert.strictEqual(actual.success, false);
  });
});
