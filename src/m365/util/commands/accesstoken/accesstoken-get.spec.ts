import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './accesstoken-get.js';

describe(commands.ACCESSTOKEN_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogRawSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    loggerLogRawSpy = sinon.spy(logger, 'logRaw');
  });

  afterEach(() => {
    sinonUtil.restore([
      auth.ensureAccessToken
    ]);
    auth.connection.accessTokens = {};
    auth.connection.spoUrl = undefined;
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ACCESSTOKEN_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves access token for the specified resource', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await command.action(logger, { options: { resource: 'https://graph.microsoft.com' } });
    assert(loggerLogSpy.calledWith('ABC'));
  });

  it('retrieves access token for SharePoint when sharepoint specified as the resource and SPO URL previously retrieved', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    auth.connection.accessTokens['https://contoso.sharepoint.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await command.action(logger, { options: { resource: 'sharepoint' } });
    assert(loggerLogSpy.calledWith('ABC'));
  });

  it('correctly handles error when retrieving access token', async () => {
    sinon.stub(auth, 'ensureAccessToken').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { resource: 'https://graph.microsoft.com' } } as any), new CommandError('An error has occurred'));
  });

  it('returns error when sharepoint specified as resource and SPO URL not available', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://contoso.sharepoint.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await assert.rejects(command.action(logger, { options: { resource: 'sharepoint' } } as any), new CommandError(`SharePoint URL undefined. Use the 'm365 spo set --url https://contoso.sharepoint.com' command to set the URL`));
  });

  it('retrieves access token for graph.microsoft.com when graph specified as the resource', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await command.action(logger, { options: { resource: 'graph' } });
    assert(loggerLogSpy.calledWith('ABC'));
  });

  it('decodes access token', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      accessToken: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNTE2MjM5MDIyfQ.SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQssw5c'
    };

    await command.action(logger, { options: { resource: 'graph', decoded: true } });
    assert(loggerLogRawSpy.calledWith(`{
  "alg": "HS256",
  "typ": "JWT"
}.{
  "sub": "1234567890",
  "name": "John Doe",
  "iat": 1516239022
}.[signature]`));
  });
});
