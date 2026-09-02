import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spe } from '../../../../utils/spe.js';
import commands from '../../commands.js';
import command, { options } from './containertype-remove.js';

describe(commands.CONTAINERTYPE_REMOVE, () => {
  const containerTypeId = 'c6f08d91-77fa-485f-9369-f246ec0fc19c';
  const containerTypeName = 'Container type name';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let confirmationPromptStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').withArgs('delegated').returns();

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
    confirmationPromptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      spe.getContainerTypeIdByName,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINERTYPE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both id and name options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: containerTypeId, name: containerTypeName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither id nor name options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ id: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if id is a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ id: containerTypeId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if name is passed', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerTypeName });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before removing the container type', async () => {
    await command.action(logger, { options: { id: containerTypeId } });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts removing the container type when prompt is not confirmed', async () => {
    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { name: containerTypeName } });
    assert(deleteStub.notCalled);
  });

  it('correctly removes a container type by id', async () => {
    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { id: containerTypeId, force: true, verbose: true } });
    assert.strictEqual(deleteStub.firstCall.args[0].url, `https://graph.microsoft.com/v1.0/storage/fileStorage/containerTypes/${containerTypeId}`);
  });

  it('correctly removes a container type by name', async () => {
    sinon.stub(spe, 'getContainerTypeIdByName').withArgs(containerTypeName).resolves(containerTypeId);
    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { name: containerTypeName, verbose: true, force: true } });
    assert.strictEqual(deleteStub.firstCall.args[0].url, `https://graph.microsoft.com/v1.0/storage/fileStorage/containerTypes/${containerTypeId}`);
  });

  it('correctly handles error when removing a container type', async () => {
    const errorMessage = 'The container type could not be deleted.';

    sinon.stub(request, 'delete').rejects({
      error: {
        code: 'badRequest',
        message: errorMessage,
        innerError: {
          date: '2026-09-02T09:00:00',
          'request-id': 'cd4a91a1-6041-c000-29c0-26f4566b5b74',
          'client-request-id': 'cd4a91a1-6041-c000-29c0-26f4566b5b74'
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { id: containerTypeId, force: true } }),
      new CommandError(errorMessage));
  });

  it('correctly handles error when retrieving a container type by name', async () => {
    const errorMessage = `The specified container type '${containerTypeName}' does not exist.`;
    sinon.stub(spe, 'getContainerTypeIdByName').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: { name: containerTypeName, force: true } }),
      new CommandError(errorMessage));
  });
});
