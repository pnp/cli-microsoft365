import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken, sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./task-checklistitem-list');

describe(commands.TASK_CHECKLISTITEM_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  
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
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
  });

  beforeEach(() => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
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
      request.get,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_CHECKLISTITEM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'isChecked']);
  });

  it('fails validation when using app only access token', (done) => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    command.action(logger, {
      options: {
        taskId: 'vzCcZoOv-U27PwydxHB8opcADJo-'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('This command does not support application permissions.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('successfully handles item found(JSON)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/vzCcZoOv-U27PwydxHB8opcADJo-/details?$select=checklist`) {
        return Promise.resolve(jsonOutput
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        taskId: 'vzCcZoOv-U27PwydxHB8opcADJo-', debug: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          jsonOutput.checklist
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('successfully handles item found(TEXT)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/vzCcZoOv-U27PwydxHB8opcADJo-/details?$select=checklist`) {
        return Promise.resolve(jsonOutput
        );
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        taskId: 'vzCcZoOv-U27PwydxHB8opcADJo-', debug: true, output: 'text'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          textOutput.checklist
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles item not found', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('The requested item is not found.'));

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The requested item is not found.')));
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
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});