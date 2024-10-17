import assert from 'assert';
import { createRequire } from 'module';
import os from 'os';
import sinon from 'sinon';
import auth, { AuthType } from '../../../Auth.js';
import { cli } from '../../../cli/cli.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command from './cli-doctor.js';

const require = createRequire(import.meta.url);
const packageJSON = require('../../../../package.json');

describe(commands.DOCTOR, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.connection.active = true;
    sinon.stub(cli.getConfig(), 'all').value({});
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
      os.platform,
      os.version,
      os.release,
      process.env
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.DOCTOR), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves scopes in the diagnostic information about the current environment', async () => {
    const jwt = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      scp: 'AllSites.FullControl AppCatalog.ReadWrite.All'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    // must be a direct assignment rather than a stub, because appId is optional
    // and undefined by default, which means it can't be stubbed
    auth.connection.appId = '31359c7f-bd7e-475c-86db-fdb8c937548e';
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.DeviceCode);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith({
      authMode: 'deviceCode',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: [],
      scopes: {
        'https://graph.microsoft.com': [
          'AllSites.FullControl',
          'AppCatalog.ReadWrite.All'
        ]
      }
    }));
  });

  it('retrieves scopes from multiple access tokens in the diagnostic information about the current environment', async () => {
    const jwt1 = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      scp: 'AllSites.FullControl AppCatalog.ReadWrite.All'
    });
    let jwt64 = Buffer.from(jwt1).toString('base64');
    const accessToken1 = `abc.${jwt64}.def`;

    const jwt2 = JSON.stringify({
      aud: 'https://mydev.sharepoint.com',
      scp: 'TermStore.Read.All'
    });
    jwt64 = Buffer.from(jwt2).toString('base64');
    const accessToken2 = `abc.${jwt64}.def`;

    const jwt3 = JSON.stringify({
      aud: 'https://mydev-admin.sharepoint.com',
      scp: 'TermStore.Read.All'
    });
    jwt64 = Buffer.from(jwt3).toString('base64');
    const accessToken3 = `abc.${jwt64}.def`;

    const jwt4 = JSON.stringify({
      aud: 'https://mydev-my.sharepoint.com',
      scp: 'TermStore.Read.All'
    });
    jwt64 = Buffer.from(jwt4).toString('base64');
    const accessToken4 = `abc.${jwt64}.def`;

    const jwt5 = JSON.stringify({
      aud: 'https://contoso-admin.sharepoint.com',
      scp: 'TermStore.Read.All'
    });
    jwt64 = Buffer.from(jwt5).toString('base64');
    const accessToken5 = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken1}` },
      'https://mydev.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken2}` },
      'https://mydev-admin.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken3}` },
      'https://mydev-my.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken4}` },
      'https://contoso-admin.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken5}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    auth.connection.appId = '31359c7f-bd7e-475c-86db-fdb8c937548e';
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.DeviceCode);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith({
      authMode: 'deviceCode',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: [],
      scopes: {
        'https://graph.microsoft.com': [
          'AllSites.FullControl',
          'AppCatalog.ReadWrite.All'
        ],
        'https://mydev.sharepoint.com': [
          'TermStore.Read.All'
        ],
        'https://contoso.sharepoint.com': [
          'TermStore.Read.All'
        ]
      }
    }));
  });

  it('retrieves roles in the diagnostic information about the current environment', async () => {
    const jwt = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      roles: ['Sites.Read.All', 'Files.ReadWrite.All']
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.DeviceCode);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith({
      authMode: 'deviceCode',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
      scopes: {}
    }));
  });

  it('retrieves roles from multiple access tokens in the diagnostic information about the current environment', async () => {
    const jwt1 = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      roles: ['Sites.Read.All', 'Files.ReadWrite.All']
    });
    let jwt64 = Buffer.from(jwt1).toString('base64');
    const accessToken1 = `abc.${jwt64}.def`;

    const jwt2 = JSON.stringify({
      aud: 'https://mydev.sharepoint.com',
      roles: ['TermStore.Read.All']
    });
    jwt64 = Buffer.from(jwt2).toString('base64');
    const accessToken2 = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken1}` },
      'https://mydev.sharepoint.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken2}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.DeviceCode);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith({
      authMode: 'deviceCode',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: ['Sites.Read.All', 'Files.ReadWrite.All', 'TermStore.Read.All'],
      scopes: {}
    }));
  });

  it('retrieves roles and scopes in the diagnostic information about the current environment', async () => {
    const jwt = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
      scp: 'Sites.Read.All Files.ReadWrite.All'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.DeviceCode);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith({
      authMode: 'deviceCode',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
      scopes: {
        'https://graph.microsoft.com':
          [
            'Sites.Read.All',
            'Files.ReadWrite.All'
          ]
      }
    }));
  });

  it('retrieves diagnostic information about the current environment when there are no roles or scopes available', async () => {
    const jwt = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      roles: [],
      scp: ''
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
    });

    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.DeviceCode);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith({
      authMode: 'deviceCode',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: [],
      scopes: {}
    }));
  });

  it('retrieves diagnostic information about the current environment with auth type Certificate', async () => {
    const jwt = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      roles: ['Sites.Read.All', 'Files.ReadWrite.All']
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.Certificate);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith({
      authMode: 'certificate',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
      scopes: {}
    }));
  });

  it('retrieves tenant information as single when TenantID is a GUID', async () => {
    const jwt = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      roles: ['Sites.Read.All', 'Files.ReadWrite.All']
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('923d42f0-6d23-41eb-b68d-c036d242654f');
    sinon.stub(auth.connection, 'authType').value(AuthType.Certificate);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith({
      authMode: 'certificate',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'single',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
      scopes: {}
    }));
  });

  it('retrieves diagnostic information about the current environment (debug)', async () => {
    const jwt = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      roles: ['Sites.Read.All', 'Files.ReadWrite.All']
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.Certificate);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith({
      authMode: 'certificate',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
      scopes: {}
    }));
  });

  it('retrieves diagnostic information of the current environment when executing in docker', async () => {
    const jwt = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      roles: ['Sites.Read.All', 'Files.ReadWrite.All']
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.Certificate);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': 'docker' });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith({
      authMode: 'certificate',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: 'docker',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: ['Sites.Read.All', 'Files.ReadWrite.All'],
      scopes: {}
    }));
  });

  it('returns empty roles and scopes in diagnostic information when access token is empty', async () => {
    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': '' }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.Certificate);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith({
      authMode: 'certificate',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: [],
      scopes: {}
    }));
  });


  it('returns empty roles and scopes in diagnostic information when access token is invalid', async () => {
    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': 'a.b.c.d' }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.Certificate);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith({
      authMode: 'certificate',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {},
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: [],
      scopes: {}
    }));
  });

  it('retrieves CLI Configuration in the diagnostic information about the current environment', async () => {
    const jwt = JSON.stringify({
      aud: 'https://graph.microsoft.com',
      scp: 'AllSites.FullControl AppCatalog.ReadWrite.All'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const accessToken = `abc.${jwt64}.def`;

    sinon.stub(auth.connection, 'accessTokens').value({
      'https://graph.microsoft.com': { 'expiresOn': '2021-07-04T09:52:18.000Z', 'accessToken': `${accessToken}` }
    });
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(os, 'version').returns('Windows 10 Pro');
    sinon.stub(os, 'release').returns('10.0.19043');
    sinon.stub(packageJSON, 'version').value('3.11.0');
    sinon.stub(process, 'version').value('v14.17.0');
    sinon.stub(auth.connection, 'appId').value('31359c7f-bd7e-475c-86db-fdb8c937548e');
    sinon.stub(auth.connection, 'tenant').value('common');
    sinon.stub(auth.connection, 'authType').value(AuthType.DeviceCode);
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinonUtil.restore(cli.getConfig().all);
    sinon.stub(cli.getConfig(), 'all').value({ "showHelpOnFailure": false });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith({
      authMode: 'deviceCode',
      cliEntraAppId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      cliEntraAppTenant: 'common',
      cliEnvironment: '',
      cliVersion: '3.11.0',
      cliConfig: {
        "showHelpOnFailure": false
      },
      nodeVersion: 'v14.17.0',
      os: { 'platform': 'win32', 'version': 'Windows 10 Pro', 'release': '10.0.19043' },
      roles: [],
      scopes: {
        'https://graph.microsoft.com': [
          'AllSites.FullControl',
          'AppCatalog.ReadWrite.All'
        ]
      }
    }));
  });
});
