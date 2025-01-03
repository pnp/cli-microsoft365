import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './container-activate.js';
import { CommandError } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.CONTAINER_ACTIVATE, () => {
  let log: string[];
  let logger: Logger;

  const containerId = 'b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_ACTIVATE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('activates container by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/activate`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: containerId, verbose: true } });
    assert(postStub.calledOnce);
  });

  it('correctly handles error when container specified by id is not found', async () => {
    const error = {
      error: {
        code: 'itemNotFound',
        message: 'Item not found',
        innerError: {
          date: '2024-10-18T09:58:41',
          'request-id': 'ec6a7cf6-4017-4af2-a3aa-82cac95dced7',
          'client-request-id': '2453c2e7-e937-52ff-1478-647fc551d4e4'
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/activate`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: containerId, verbose: true } } as any),
      new CommandError(error.error.message));
  });
});