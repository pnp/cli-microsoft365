import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { odata } from '../../../../utils/odata';
import { Cli } from '../../../../cli/Cli';
const command: Command = require('./user-recyclebinitem-clear');

describe(commands.USER_RECYCLEBINITEM_CLEAR, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;

  const deletedUsersResponse = [{ id: '4c099956-ca9a-4e60-ad5f-3f8447122706' }];
  const graphGetUrl = 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?$select=id';
  const graphBatchUrl = 'https://graph.microsoft.com/v1.0/$batch';

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      odata.getAllItems,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
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

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

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

    await command.action(logger, { options: { confirm: true, verbose: true } });
    assert.strictEqual(amountOfBatches, 3);
  });

  it('prompts before removing users', async () => {
    await command.action(logger, { options: {} });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('aborts removing users when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    const postStub = sinon.stub(request, 'post').callsFake(async () => {
      return;
    });
    await command.action(logger, { options: {} });
    assert(postStub.notCalled);
  });

  it('correctly handles API error', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async () => {
      throw {
        error: {
          error: {
            code: 'Invalid_Request',
            message: 'An error has occured while processing this request.',
            innerError: {
              'request-id': '9b0df954-93b5-4de9-8b99-43c204a8aaf8',
              date: '2018-04-24T18:56:48'
            }
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { confirm: true } } as any),
      new CommandError('An error has occured while processing this request.'));
  });
});
