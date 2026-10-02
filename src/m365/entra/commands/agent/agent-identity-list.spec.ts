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
import command, { options } from './agent-identity-list.js';

describe(commands.AGENT_IDENTITY_LIST, () => {
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const identities = [{
    id: '1b7313c4-05d0-4a08-88e3-7b76c003a0a2',
    displayName: 'My Agent Identity',
    createdDateTime: '2019-09-17T19:10:35Z',
    createdByAppId: '00001111-aaaa-2222-bbbb-3333cccc4444',
    agentIdentityBlueprintId: '00001111-aaaa-2222-bbbb-3333cccc4444',
    accountEnabled: true,
    disabledByMicrosoftStatus: null,
    servicePrincipalType: 'ServiceIdentity',
    tags: []
  }];
  const propertiesResponse = [{
    displayName: 'My Agent Identity',
    id: '1b7313c4-05d0-4a08-88e3-7b76c003a0a2'
  }];

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
    logger = {
      log: async () => {},
      logRaw: async () => {},
      logToStderr: async () => {}
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => sinonUtil.restore([request.get]));
  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.AGENT_IDENTITY_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName']);
  });

  it('lists agent identities', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      assert.strictEqual(opts.url, 'https://graph.microsoft.com/v1.0/servicePrincipals/microsoft.graph.agentIdentity');
      assert.strictEqual(opts.responseType, 'json');
      return { value: identities };
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert(loggerLogSpy.calledWith(identities));
  });

  it('lists agent identities with the specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      assert.strictEqual(opts.url, 'https://graph.microsoft.com/v1.0/servicePrincipals/microsoft.graph.agentIdentity?$select=id%2CdisplayName');
      return { value: propertiesResponse };
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ properties: 'id,displayName' }) });
    assert(loggerLogSpy.calledWith(propertiesResponse));
  });

  it('removes quotes and whitespace from the specified properties and ignores empty values', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      assert.strictEqual(opts.url, 'https://graph.microsoft.com/v1.0/servicePrincipals/microsoft.graph.agentIdentity?$select=id%2CdisplayName');
      return { value: propertiesResponse };
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ properties: "'id' , 'displayName'," }) });
    assert(loggerLogSpy.calledWith(propertiesResponse));
  });

  it('escapes the specified properties so that they cannot inject additional query parameters', async () => {
    let requestedUrl = '';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      requestedUrl = opts.url!;
      return { value: propertiesResponse };
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ properties: 'id&$filter=displayName eq \'My Agent Identity\'' }) });

    const parsedUrl = new URL(requestedUrl);
    assert.strictEqual(parsedUrl.searchParams.get('$select'), 'id&$filter=displayName eq My Agent Identity');
    assert.strictEqual(parsedUrl.searchParams.get('$filter'), null);
    assert.strictEqual(parsedUrl.hash, '');
    assert.strictEqual([...parsedUrl.searchParams.keys()].length, 1);
  });

  it('escapes fragments in the specified properties so that they cannot truncate the query string', async () => {
    let requestedUrl = '';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      requestedUrl = opts.url!;
      return { value: propertiesResponse };
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ properties: 'id#displayName' }) });

    const parsedUrl = new URL(requestedUrl);
    assert.strictEqual(parsedUrl.searchParams.get('$select'), 'id#displayName');
    assert.strictEqual(parsedUrl.hash, '');
    assert.strictEqual([...parsedUrl.searchParams.keys()].length, 1);
  });

  it('does not add a select query parameter when no top-level property is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      assert.strictEqual(opts.url, 'https://graph.microsoft.com/v1.0/servicePrincipals/microsoft.graph.agentIdentity');
      return { value: identities };
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ properties: 'manager/displayName' }) });
    assert(loggerLogSpy.calledWith(identities));
  });

  it('handles error when retrieving the agent identities list failed', async () => {
    sinon.stub(request, 'get').rejects({ error: { message: 'Permission denied' } });
    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('Permission denied'));
  });
});
