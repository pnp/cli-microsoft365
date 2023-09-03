import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './group-list.js';

describe(commands.GROUP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'groupType']);
  });

  it('lists aad Groups in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups`) {
        return {
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
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
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
  });

  it('lists aad Groups in the tenant (text)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups`) {
        return {
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
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, output: 'text' } });
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
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: {} }), new CommandError(errorMessage));
  });
});
