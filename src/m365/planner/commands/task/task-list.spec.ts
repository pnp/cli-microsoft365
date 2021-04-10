import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./task-list');

describe(commands.TASK_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    assert.strictEqual(command.name.startsWith(commands.TASK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'startDateTime', 'dueDateTime', 'completedDateTime']);
  });

  it('lists planner tasks of the current logged in user', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/planner/tasks`) {
        return Promise.resolve({
          "value": [
            {
              "planId": "IlGTfsb-PEWl5EYIx97I5WUAB8ni",
              "bucketId": "fno1rNw2Vk2x7XwLQib9aWUAC2YS",
              "title": "Northwind HR Training Video Part I",
              "orderHint": "8586967557616915534",
              "assigneePriority": "8586967557616915534",
              "percentComplete": 50,
              "startDateTime": "2017-09-08T00:00:00Z",
              "createdDateTime": "2017-09-08T06:12:03.7860273Z",
              "dueDateTime": "2018-09-03T00:00:00Z",
              "hasDescription": false,
              "previewType": "description",
              "completedDateTime": null,
              "completedBy": null,
              "referenceCount": 1,
              "checklistItemCount": 0,
              "activeChecklistItemCount": 0,
              "conversationThreadId": null,
              "id": "102sl-tTCkyFHptTaFW5lGUACsAe",
              "createdBy": {
                "user": {
                  "displayName": null,
                  "id": "48d31887-5fad-4d73-a9f5-3c356e68a038"
                }
              },
              "appliedCategories": {},
              "assignments": {
                "48d31887-5fad-4d73-a9f5-3c356e68a038": {
                  "@odata.type": "#microsoft.graph.plannerAssignment",
                  "assignedDateTime": "2017-09-08T06:12:03.7860273Z",
                  "orderHint": "",
                  "assignedBy": {
                    "user": {
                      "displayName": null,
                      "id": "48d31887-5fad-4d73-a9f5-3c356e68a038"
                    }
                  }
                }
              }
            },
            {
              "planId": "Ey4oAJeTv0W6kx-kD4T-kGUAHEwE",
              "bucketId": "XxJ8fhM6gE-2-ShejgmMWGUAEVtB",
              "title": "Search Optimization",
              "orderHint": "8586967558658533417",
              "assigneePriority": "8586967558658533417",
              "percentComplete": 0,
              "startDateTime": "2017-09-03T00:00:00Z",
              "createdDateTime": "2017-09-08T06:10:19.624239Z",
              "dueDateTime": "2018-08-29T00:00:00Z",
              "hasDescription": false,
              "previewType": "automatic",
              "completedDateTime": null,
              "completedBy": null,
              "referenceCount": 0,
              "checklistItemCount": 0,
              "activeChecklistItemCount": 0,
              "conversationThreadId": null,
              "id": "7aZeJUYK90OZiFq6H7Ug3mUACcdr",
              "createdBy": {
                "user": {
                  "displayName": null,
                  "id": "08fa38e4-cbfa-4488-94ed-c834da6539df"
                }
              },
              "appliedCategories": {},
              "assignments": {
                "48d31887-5fad-4d73-a9f5-3c356e68a038": {
                  "@odata.type": "#microsoft.graph.plannerAssignment",
                  "assignedDateTime": "2017-09-08T06:10:19.624239Z",
                  "orderHint": "",
                  "assignedBy": {
                    "user": {
                      "displayName": null,
                      "id": "08fa38e4-cbfa-4488-94ed-c834da6539df"
                    }
                  }
                }
              }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(loggerLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
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