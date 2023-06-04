import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./schemaextension-remove');

describe(commands.SCHEMAEXTENSION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptOptions: any;

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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
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

    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension', confirm: true } });
    assert(loggerLogSpy.notCalled);
  });

  it('removes schema extension (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/schemaExtensions/`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 'exttyee4dv5_MySchemaExtension', confirm: true } });
    assert(loggerLogToStderrSpy.called);
  });

  it('prompts before removing schema extension when confirmation argument not passed', async () => {
    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing schema extension when prompt not confirmed', async () => {
    sinon.stub(request, 'delete').rejects('Invalid request');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: false });

    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension' } });
  });

  it('removes schema extension when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`v1.0/schemaExtensions/`) > -1) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'delete').rejects({ error: 'An error has occurred' });

    await assert.rejects(command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension', confirm: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles random API error (string error)', async () => {
    sinon.stub(request, 'delete').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { id: 'exttyee4dv5_MySchemaExtension', confirm: true } } as any),
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
