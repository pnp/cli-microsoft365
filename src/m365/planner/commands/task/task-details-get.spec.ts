import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./task-details-get');

describe(commands.TASK_DETAILS_GET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TASK_DETAILS_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('successfully handles item found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/vzCcZoOv-U27PwydxHB8opcADJo-/details`) {
        return Promise.resolve({
          "description": "Description",
          "previewType": "checklist",
          "id": "vzCcZoOv-U27PwydxHB8opcADJo-",
          "references": {
            "https%3A//www%2Econtoso%2Ecom": {
              "alias": "Contoso.com",
              "type": "Other",
              "previewPriority": "8585576049615477185P<",
              "lastModifiedDateTime": "2022-02-04T19:13:03.9611197Z",
              "lastModifiedBy": {
                "user": {
                  "displayName": null,
                  "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
                }
              }
            }
          },
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
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        taskId: 'vzCcZoOv-U27PwydxHB8opcADJo-', debug: true
      }
    }, () => {
      try {
        const actual = JSON.stringify(log[log.length - 1]);
        const expected = JSON.stringify({
          "description": "Description",
          "previewType": "checklist",
          "id": "vzCcZoOv-U27PwydxHB8opcADJo-",
          "references": {
            "https%3A//www%2Econtoso%2Ecom": {
              "alias": "Contoso.com",
              "type": "Other",
              "previewPriority": "8585576049615477185P<",
              "lastModifiedDateTime": "2022-02-04T19:13:03.9611197Z",
              "lastModifiedBy": {
                "user": {
                  "displayName": null,
                  "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
                }
              }
            }
          },
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
