import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './schemaextension-remove.js';

describe(commands.SCHEMAEXTENSION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptIssued: boolean = false;

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SCHEMAEXTENSION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes schema extension', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/schemaExtensions/`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension', force: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('removes schema extension (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/schemaExtensions/`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'exttyee4dv5_MySchemaExtension', force: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('prompts before removing schema extension when confirmation argument not passed', async () => {
    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension' } });

    assert(promptIssued);
  });

  it('aborts removing schema extension when prompt not confirmed', async () => {
    sinon.stub(request, 'delete').rejects('Invalid request');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension' } });
  });

  it('removes schema extension when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`v1.0/schemaExtensions/`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'delete').rejects({ error: 'An error has occurred' });

    await assert.rejects(command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (string error)', async () => {
    sinon.stub(request, 'delete').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension', force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
