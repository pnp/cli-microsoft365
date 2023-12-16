import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { odata } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-recyclebinitem-clear.js';

describe(commands.USER_RECYCLEBINITEM_CLEAR, () => {
  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;

  const deletedUsersResponse = [{ id: '4c099956-ca9a-4e60-ad5f-3f8447122706' }];
  const graphGetUrl = 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?$select=id';
  const graphBatchUrl = 'https://graph.microsoft.com/v1.0/$batch';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      odata.getAllItems,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_RECYCLEBINITEM_CLEAR);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes a single user when prompt confirmed', async () => {
    let amountOfBatches = 0;

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === graphGetUrl) {
        return deletedUsersResponse;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === graphBatchUrl) {
        amountOfBatches++;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert.strictEqual(amountOfBatches, 1);
  });

  it('removes users using multiple batches', async () => {
    let amountOfBatches = 0;
    const deletedUsersResponseLarge: any[] = [];
    for (let i = 0; i < 50; i++) {
      deletedUsersResponseLarge.push(deletedUsersResponse[0]);
    }
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === graphGetUrl) {
        return deletedUsersResponseLarge;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === graphBatchUrl) {
        amountOfBatches++;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { force: true, verbose: true } });
    assert.strictEqual(amountOfBatches, 3);
  });

  it('prompts before removing users', async () => {
    await command.action(logger, { options: {} });
    assert(promptIssued);
  });

  it('aborts removing users when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    const postStub = sinon.stub(request, 'post').callsFake(async () => {
      return;
    });

    await command.action(logger, { options: {} });
    assert(postStub.notCalled);
  });

  it('correctly handles API error', async () => {
    sinon.stub(odata, 'getAllItems').rejects({
      error: {
        error: {
          code: 'Invalid_Request',
          message: 'An error has occurred while processing this request.',
          innerError: {
            'request-id': '9b0df954-93b5-4de9-8b99-43c204a8aaf8',
            date: '2018-04-24T18:56:48'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { force: true } } as any),
      new CommandError('An error has occurred while processing this request.'));
  });
});
