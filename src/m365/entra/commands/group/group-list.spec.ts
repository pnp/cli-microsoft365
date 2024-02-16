import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './group-list.js';
import aadCommands from '../../aadCommands.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { settingsNames } from '../../../../settingsNames.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.GROUP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const userId = '00000000-0000-0000-0000-000000000000';
  const userName = 'john@contoso.com';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(cli, 'getSettingWithDefaultValue').returnsArg(1);
    auth.connection.active = true;
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        accessToken: 'abc'
      };
    }
    commandInfo = cli.getCommandInfo(command);
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
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.GROUP_LIST]);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'groupType']);
  });

  it('fails validation when invalid type specified', async () => {
    const actual = await command.validate({
      options: {
        type: 'Invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid type specified', async () => {
    const actual = await command.validate({
      options: {
        type: 'microsoft365'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { joined: true, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { joined: true, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation userName is used without joined or associated option', async () => {
    const actual = await command.validate({ options: { userName: userName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation userId is used without joined or associated option', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both userId and userName are used', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { joined: true, userId: userId, userName: userName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both joined and associated options are used', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { joined: true, associated: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('lists all entra groups in the tenant (verbose)', async () => {
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
    assert(loggerLogSpy.calledOnceWithExactly([
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

  it('lists all entra groups in the tenant (text)', async () => {
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
    assert(loggerLogSpy.calledOnceWithExactly([
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

  it('lists all microsoft365 groups in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')`) {
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
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'microsoft365' } });
    assert(loggerLogSpy.calledOnceWithExactly([
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
      }
    ]));
  });

  it('lists all distribution groups in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=securityEnabled eq false and mailEnabled eq true&$count=true`) {
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
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'distribution' } });
    assert(loggerLogSpy.calledOnceWithExactly([
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
      }
    ]));
  });

  it('lists all security groups in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=securityEnabled eq true and mailEnabled eq false&$count=true`) {
        return {
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "groupTypes": [],
              "mail": null,
              "mailEnabled": false,
              "mailNickname": "00000000-0000-0000-0000-000000000000",
              "securityEnabled": true
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'security' } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
        "description": "Code Challenge",
        "displayName": "Code Challenge",
        "groupTypes": [],
        "mail": null,
        "mailEnabled": false,
        "mailNickname": "00000000-0000-0000-0000-000000000000",
        "securityEnabled": true
      }
    ]));
  });

  it('lists all mailEnabledSecurity groups in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=securityEnabled eq true and mailEnabled eq true and not(groupTypes/any(t:t eq 'Unified'))&$count=true`) {
        return {
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "groupTypes": [],
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "securityEnabled": true
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'mailEnabledSecurity' } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
        "description": "Code Challenge",
        "displayName": "Code Challenge",
        "groupTypes": [],
        "mail": "CodeChallenge@dev1802.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "CodeChallenge",
        "securityEnabled": true
      }
    ]));
  });

  it('lists all groups in a tenant that the currently signed in user is a part of', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/memberOf/microsoft.graph.group`) {
        return {
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "groupTypes": [],
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "securityEnabled": true
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { joined: true } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
        "description": "Code Challenge",
        "displayName": "Code Challenge",
        "groupTypes": [],
        "mail": "CodeChallenge@dev1802.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "CodeChallenge",
        "securityEnabled": true
      }
    ]));
  });

  it('lists all groups in a tenant that the currently signed in user is associated with', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.group`) {
        return {
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "groupTypes": [],
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "securityEnabled": true
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { associated: true } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
        "description": "Code Challenge",
        "displayName": "Code Challenge",
        "groupTypes": [],
        "mail": "CodeChallenge@dev1802.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "CodeChallenge",
        "securityEnabled": true
      }
    ]));
  });

  it('lists all groups in a tenant user specified by id is associated with', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/transitiveMemberOf/microsoft.graph.group`) {
        return {
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "groupTypes": [],
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "securityEnabled": true
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId, associated: true } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
        "description": "Code Challenge",
        "displayName": "Code Challenge",
        "groupTypes": [],
        "mail": "CodeChallenge@dev1802.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "CodeChallenge",
        "securityEnabled": true
      }
    ]));
  });

  it('lists all groups in a tenant user specified by id is a part of', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName)}/memberOf/microsoft.graph.group`) {
        return {
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "groupTypes": [],
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "securityEnabled": true
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName, joined: true } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
        "description": "Code Challenge",
        "displayName": "Code Challenge",
        "groupTypes": [],
        "mail": "CodeChallenge@dev1802.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "CodeChallenge",
        "securityEnabled": true
      }
    ]));
  });

  it('throws error when retrieving all joined groups for the current logged in user when using application permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        joined: true
      }
    }), new CommandError('You must specify either userId or userName when using application only permissions and specifying the joined option'));
  });

  it('throws error when retrieving all associated groups for the current logged in user when using application permissions', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        associated: true
      }
    }), new CommandError('You must specify either userId or userName when using application only permissions and specifying the associated option'));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: {} }), new CommandError(errorMessage));
  });
});
