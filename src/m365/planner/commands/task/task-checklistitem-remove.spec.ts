import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./task-checklistitem-remove');

describe(commands.TASK_CHECKLISTITEM_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validId = '71175';

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    sinon.stub(Cli.getInstance().config, 'all').value({});
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
    promptOptions = undefined;

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: true });
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      Cli.getInstance().config.all
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_CHECKLISTITEM_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removal when confirm option not passed', (done) => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });

    command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId
      }
    }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation when valid options specified', (done) => {
    const actual = command.validate({
      options: {
        taskId: validTaskId,
        id: validId
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('Correctly deletes checklist item', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details`) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }
      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details`) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid Request');
    });
    command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId,
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when Planner task does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details`) {
        return Promise.reject('The request item is not found.');
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Planner task was not found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});