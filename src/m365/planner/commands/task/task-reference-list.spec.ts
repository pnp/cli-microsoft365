import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./task-reference-list');

describe(commands.TASK_REFERENCE_LIST, () => {
  const referenceListResponse = {
    "https%3A//contoso%2Esharepoint%2Ecom/sites/HRPlan/Shared Documents/Sample.pdf": {
      "alias": "Sample.pdf",
      "type": "Pdf",
      "previewPriority": "[>",
      "lastModifiedDateTime": "2022-05-15T16:20:31.8649232Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "fe36f75f-c103-410b-a18a-2bf6df06ac3a"
        }
      }
    },
    "https%3A//contoso%2Esharepoint%2Ecom/sites/HRPlan/Shared Documents/Sample.png": {
      "alias": "Sample.png",
      "type": "Other",
      "previewPriority": "8585492445655664725P(",
      "lastModifiedDateTime": "2022-05-12T13:32:59.9267487Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "fe36f75f-c103-410b-a18a-2bf6df06ac3a"
        }
      }
    }
  };
  
  
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_REFERENCE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('successfully handles item found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/uBk5fK_MHkeyuPYlCo4OFpcAMowf/details?$select=references`) {
        return Promise.resolve(referenceListResponse);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        taskId: 'uBk5fK_MHkeyuPYlCo4OFpcAMowf', debug: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(referenceListResponse));
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
