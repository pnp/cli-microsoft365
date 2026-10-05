import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './accesstoken-get.js';

describe(commands.ACCESSTOKEN_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogRawSpy: sinon.SinonSpy;
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

    await command.action(logger, { options: commandOptionsSchema.parse({ resource: 'https://graph.microsoft.com' }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ resource: 'sharepoint' }) });
    assert(loggerLogSpy.calledWith('ABC'));
  });

  it('correctly handles error when retrieving access token', async () => {
    sinon.stub(auth, 'ensureAccessToken').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ resource: 'https://graph.microsoft.com' }) }), new CommandError('An error has occurred'));
  });

  it('returns error when sharepoint specified as resource and SPO URL not available', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://contoso.sharepoint.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ resource: 'sharepoint' }) }), new CommandError(`SharePoint URL undefined. Use the 'm365 spo set --url https://contoso.sharepoint.com' command to set the URL`));
  });

  it('retrieves access token for graph.microsoft.com when graph specified as the resource', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await command.action(logger, { options: commandOptionsSchema.parse({ resource: 'graph' }) });
    assert(loggerLogSpy.calledWith('ABC'));
  });

  it('decodes access token', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      accessToken: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNTE2MjM5MDIyfQ.SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQssw5c'
    };

    await command.action(logger, { options: commandOptionsSchema.parse({ resource: 'graph', decoded: true }) });
    assert(loggerLogRawSpy.calledWith(`{
  "alg": "HS256",
  "typ": "JWT"
}.{
  "sub": "1234567890",
  "name": "John Doe",
  "iat": 1516239022
}.[signature]`));
  });

  it('fails validation if resource is missing', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      resource: 'https://graph.microsoft.com',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with all required options', () => {
    const actual = commandOptionsSchema.safeParse({ resource: 'https://graph.microsoft.com' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with optional new flag', () => {
    const actual = commandOptionsSchema.safeParse({ resource: 'https://graph.microsoft.com', new: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with optional decoded flag', () => {
    const actual = commandOptionsSchema.safeParse({ resource: 'https://graph.microsoft.com', decoded: true });
    assert.strictEqual(actual.success, true);
  });
});
