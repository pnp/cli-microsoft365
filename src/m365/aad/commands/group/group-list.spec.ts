import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./group-list');

describe(commands.GROUP_LIST, () => {
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
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'groupType']);
  });

  it('lists aad Groups in the tenant (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups`) {
        return Promise.resolve({
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "groupTypes": [
                "Unified"
              ],
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "securityEnabled": false
            },
            {
              "id": "2f64f70d-386b-489f-805a-670cad739fde",
              "description": "The Jumping Jacks",
              "displayName": "The Jumping Jacks",
              "groupTypes": [
              ],
              "mail": "TheJumpingJacks@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "TheJumpingJacks",
              "securityEnabled": true
            },
            {
              "id": "ff0554cc-8aa8-40f2-a369-ed604503fb79",
              "description": "Emergency Response",
              "displayName": "Emergency Response",
              "groupTypes": [
              ],
              "mail": null,
              "mailEnabled": false,
              "mailNickname": "00000000-0000-0000-0000-000000000000",
              "securityEnabled": true
            },
            {
              "id": "0a0bf25a-2de0-40de-9908-c96941a2615b",
              "description": "Free Birds",
              "displayName": "Free Birds",
              "groupTypes": [
              ],
              "mail": "FreeBirds@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "FreeBirds",
              "securityEnabled": false
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { verbose: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
            "description": "Code Challenge",
            "displayName": "Code Challenge",
            "groupTypes": [
              "Unified"
            ],
            "mail": "CodeChallenge@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "CodeChallenge",
            "securityEnabled": false
          },
          {
            "id": "2f64f70d-386b-489f-805a-670cad739fde",
            "description": "The Jumping Jacks",
            "displayName": "The Jumping Jacks",
            "groupTypes": [
            ],
            "mail": "TheJumpingJacks@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "TheJumpingJacks",
            "securityEnabled": true
          },
          {
            "id": "ff0554cc-8aa8-40f2-a369-ed604503fb79",
            "description": "Emergency Response",
            "displayName": "Emergency Response",
            "groupTypes": [
            ],
            "mail": null,
            "mailEnabled": false,
            "mailNickname": "00000000-0000-0000-0000-000000000000",
            "securityEnabled": true
          },
          {
            "id": "0a0bf25a-2de0-40de-9908-c96941a2615b",
            "description": "Free Birds",
            "displayName": "Free Birds",
            "groupTypes": [
            ],
            "mail": "FreeBirds@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "FreeBirds",
            "securityEnabled": false
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists deleted groups in the tenant with the default properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group`) {
        return Promise.resolve({
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "groupTypes": [
                "Unified"
              ],
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "securityEnabled": false
            },
            {
              "id": "2f64f70d-386b-489f-805a-670cad739fde",
              "description": "The Jumping Jacks",
              "displayName": "The Jumping Jacks",
              "groupTypes": [
              ],
              "mail": "TheJumpingJacks@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "TheJumpingJacks",
              "securityEnabled": true
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, deleted: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
            "description": "Code Challenge",
            "displayName": "Code Challenge",
            "groupTypes": [
              "Unified"
            ],
            "mail": "CodeChallenge@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "CodeChallenge",
            "securityEnabled": false
          },
          {
            "id": "2f64f70d-386b-489f-805a-670cad739fde",
            "description": "The Jumping Jacks",
            "displayName": "The Jumping Jacks",
            "groupTypes": [
            ],
            "mail": "TheJumpingJacks@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "TheJumpingJacks",
            "securityEnabled": true
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists aad Groups in the tenant (text)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups`) {
        return Promise.resolve({
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "groupTypes": [
                "Unified"
              ],
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "securityEnabled": false
            },
            {
              "id": "2f64f70d-386b-489f-805a-670cad739fde",
              "description": "The Jumping Jacks",
              "displayName": "The Jumping Jacks",
              "groupTypes": [
              ],
              "mail": "TheJumpingJacks@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "TheJumpingJacks",
              "securityEnabled": true
            },
            {
              "id": "ff0554cc-8aa8-40f2-a369-ed604503fb79",
              "description": "Emergency Response",
              "displayName": "Emergency Response",
              "groupTypes": [
              ],
              "mail": null,
              "mailEnabled": false,
              "mailNickname": "00000000-0000-0000-0000-000000000000",
              "securityEnabled": true
            },
            {
              "id": "0a0bf25a-2de0-40de-9908-c96941a2615b",
              "description": "Free Birds",
              "displayName": "Free Birds",
              "groupTypes": [
              ],
              "mail": "FreeBirds@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "FreeBirds",
              "securityEnabled": false
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, output: 'text' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
            "description": "Code Challenge",
            "displayName": "Code Challenge",
            "groupTypes": [
              "Unified"
            ],
            "mail": "CodeChallenge@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "CodeChallenge",
            "securityEnabled": false,
            "groupType": "Microsoft 365"
          },
          {
            "id": "2f64f70d-386b-489f-805a-670cad739fde",
            "description": "The Jumping Jacks",
            "displayName": "The Jumping Jacks",
            "groupTypes": [
            ],
            "mail": "TheJumpingJacks@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "TheJumpingJacks",
            "securityEnabled": true,
            "groupType": "Mail enabled security"
          },
          {
            "id": "ff0554cc-8aa8-40f2-a369-ed604503fb79",
            "description": "Emergency Response",
            "displayName": "Emergency Response",
            "groupTypes": [
            ],
            "mail": null,
            "mailEnabled": false,
            "mailNickname": "00000000-0000-0000-0000-000000000000",
            "securityEnabled": true,
            "groupType": "Security"
          },
          {
            "id": "0a0bf25a-2de0-40de-9908-c96941a2615b",
            "description": "Free Birds",
            "displayName": "Free Birds",
            "groupTypes": [
            ],
            "mail": "FreeBirds@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "FreeBirds",
            "securityEnabled": false,
            "groupType": "Distribution"
          }
        ]));
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
