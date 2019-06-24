import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./planner-task-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.PLANNER_TASK_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.service = new Service();
    telemetry = null;
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.PLANNER_TASK_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.PLANNER_TASK_LIST);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to Microsoft Graph', (done) => {
    auth.service = new Service();
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to the Microsoft Graph first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists planner tasks of the currnet logged in user as a JSON result', (done) => {
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

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
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

  it('lists planner tasks of the currnet logged in user', (done) => {
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

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
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

  it('lists planner tasks of the currnet logged in user (debug)', (done) => {
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

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.PLANNER_TASK_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});