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
  let loggerLogSpy: sinon.SinonSpy;
  let promptOptions: any;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validId = '71175';
  const jsonOutput = {
    "checklist": {
      "33224": {
        "isChecked": false,
        "title": "Some checklist",
        "orderHint": "8585576049720396756P(",
        "lastModifiedDateTime": "2022-02-04T19:12:53.4692149Z",
        "lastModifiedBy": {
          "user": {
            "displayName": null,
            "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
          }
        }
      },
      "69115": {
        "isChecked": false,
        "title": "Some checklist more",
        "orderHint": "85855760494@",
        "lastModifiedDateTime": "2022-02-04T19:12:55.4735671Z",
        "lastModifiedBy": {
          "user": {
            "displayName": null,
            "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
          }
        }
      }
    }
  };
  const textOutput = {
    "checklist": [{
      "id": "33224",
      "isChecked": false,
      "title": "Some checklist",
      "orderHint": "8585576049720396756P(",
      "lastModifiedDateTime": "2022-02-04T19:12:53.4692149Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
        }
      }
    },
    {
      "id": "69115",
      "isChecked": false,
      "title": "Some checklist more",
      "orderHint": "85855760494@",
      "lastModifiedDateTime": "2022-02-04T19:12:55.4735671Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
        }
      }
    }
    ]
  };


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
    loggerLogSpy = sinon.spy(logger, 'log');

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

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'isChecked']);
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

  it('correctly removes checklistitem', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details`) {
        return Promise.resolve(jsonOutput);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      taskId: validTaskId,
      id: validId,
      confirm: true
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(jsonOutput.checklist));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('correctly removes checklistitem(text)', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details`) {
        return Promise.resolve(jsonOutput);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      taskId: validTaskId,
      id: validId,
      confirm: true,
      output: 'text'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(textOutput.checklist));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when task details endpoint fails', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('invalid')}/details`) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        taskId: 'invalid',
        id: validId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Error fetching task details`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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