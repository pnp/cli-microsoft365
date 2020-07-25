import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./task-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.PLANNER_TASK_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.PLANNER_TASK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists planner tasks of the current logged in user as a JSON result', (done) => {
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

    cmdInstance.action({ options: { debug: true, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists planner tasks of the current logged in user (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/planner/tasks`) {
        return Promise.resolve({
          "value": [
            {
              "@odata.etag": "W/\"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAXCc=\"",
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
              "@odata.etag": "W/\"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAWCc=\"",
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

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});