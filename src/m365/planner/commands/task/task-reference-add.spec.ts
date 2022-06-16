import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./task-reference-add');

describe(commands.TASK_REFERENCE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validUrl = 'https://www.microsoft.com';
  const validAlias = 'Test';
  const validType = 'Word';

  const referenceResponse = {
    "https%3A//www%2Emicrosoft%2Ecom": {
      "alias": "Test",
      "type": "Word",
      "previewPriority": "8585493318091789098Pa",
      "lastModifiedDateTime": "2022-05-11T13:18:56.3142944Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "dd8b99a7-77c6-4238-a609-396d27844921"
        }
      }
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_REFERENCE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if incorrect type is specified.', (done) => {
    const actual = command.validate({
      options: {
        taskId: validTaskId,
        url: validUrl,
        type: "wrong"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid options specified', (done) => {
    const actual = command.validate({
      options: {
        taskId: validTaskId,
        url: validUrl
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly adds reference', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details`) {
        return Promise.resolve({
          references: referenceResponse
        });
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
      url: validUrl
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(referenceResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly adds reference with type and alias', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}/details`) {
        return Promise.resolve({
          references: referenceResponse
        });
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
      url: validUrl,
      alias: validAlias,
      type: validType
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(referenceResponse));
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
        url: validUrl
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