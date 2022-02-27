import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./task-get');

describe(commands.TASK_GET, () => {
  let log: string[];
  let logger: Logger;

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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('successfully handles item found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/tasks/01gzSlKkIUSUl6DF_EilrmQAKDhh`) {
        return Promise.resolve({
          "createdBy": {
            "user": {
              "id": "6463a5ce-2119-4198-9f2a-628761df4a62"
            }
          },
          "planId": "xqQg5FS2LkCp935s-FIFm2QAFkHM",
          "bucketId": "gcrYAaAkgU2EQUvpkNNXLGQAGTtu",
          "title": "title-value",
          "orderHint": "9223370609546166567W",
          "assigneePriority": "90057581\"",
          "createdDateTime": "2015-03-25T18:36:49.2407981Z",
          "assignments": {
            "fbab97d0-4932-4511-b675-204639209557": {
              "@odata.type": "#microsoft.graph.plannerAssignment",
              "assignedBy": {
                "user": {
                  "id": "1e9955d2-6acd-45bf-86d3-b546fdc795eb"
                }
              },
              "assignedDateTime": "2015-03-25T18:38:21.956Z",
              "orderHint": "RWk1"
            }
          },
          "priority": 5,
          "id": "01gzSlKkIUSUl6DF_EilrmQAKDhh"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        id: '01gzSlKkIUSUl6DF_EilrmQAKDhh', debug: true
      }
    }, () => {
      try {
        const actual = JSON.stringify(log[log.length - 1]);
        const expected = JSON.stringify({
          "createdBy": {
            "user": {
              "id": "6463a5ce-2119-4198-9f2a-628761df4a62"
            }
          },
          "planId": "xqQg5FS2LkCp935s-FIFm2QAFkHM",
          "bucketId": "gcrYAaAkgU2EQUvpkNNXLGQAGTtu",
          "title": "title-value",
          "orderHint": "9223370609546166567W",
          "assigneePriority": "90057581\"",
          "createdDateTime": "2015-03-25T18:36:49.2407981Z",
          "assignments": {
            "fbab97d0-4932-4511-b675-204639209557": {
              "@odata.type": "#microsoft.graph.plannerAssignment",
              "assignedBy": {
                "user": {
                  "id": "1e9955d2-6acd-45bf-86d3-b546fdc795eb"
                }
              },
              "assignedDateTime": "2015-03-25T18:38:21.956Z",
              "orderHint": "RWk1"
            }
          },
          "priority": 5,
          "id": "01gzSlKkIUSUl6DF_EilrmQAKDhh"
        });
        assert.strictEqual(actual, expected);

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles item not found', (done) => {
    Utils.restore(request.get);
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
    Utils.restore(request.get);
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