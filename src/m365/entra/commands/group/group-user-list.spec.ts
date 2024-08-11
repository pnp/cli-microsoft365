import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { settingsNames } from '../../../../settingsNames.js';
import { formatting } from '../../../../utils/formatting.js';
import commands from '../../commands.js';
import command from './group-user-list.js';
import aadCommands from '../../aadCommands.js';

describe(commands.GROUP_USER_LIST, () => {
  const groupId = '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a';
  const groupName = 'CLI Test Group';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
      entraGroup.getGroupIdByDisplayName,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_USER_LIST);
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
    assert.deepStrictEqual(alias, [aadCommands.GROUP_USER_LIST]);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'userPrincipalName', 'roles']);
  });

  it('fails validation if the groupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        groupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        role: 'Invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid groupId and no role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: groupId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly lists all users in a Microsoft Entra group by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "givenName": "Karl", "surname": "Matteson" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, groupId: groupId } });

    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Owner", "Member"]
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "givenName": "Karl",
        "surname": "Matteson",
        "roles": ["Member"]
      }
    ]));
  });

  it('correctly lists all users in a Microsoft Entra group by name', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves(groupId);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "givenName": "Karl", "surname": "Matteson" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, groupName: groupName } });

    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Owner", "Member"]
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "givenName": "Karl",
        "surname": "Matteson",
        "roles": ["Member"]
      }
    ]));
  });

  it('correctly lists all owners in a Microsoft Entra group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews" }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, role: "Owner" } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Owner"]
      }
    ]));
  });

  it('handles error when multiple Microsoft Entra groups with the specified displayName found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(groupName)}'&$select=id`) {
        return {
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        groupName: groupName
      }
    }), new CommandError(`Multiple groups with name 'CLI Test Group' found. Found: 9b1b1e42-794b-4c71-93ac-5ed92488b67f, 9b1b1e42-794b-4c71-93ac-5ed92488b67g.`));
  });

  it('handles selecting single result when multiple Microsoft Entra groups with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(groupName)}'&$select=id`) {
        return {
          value: [
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' },
            { id: '9b1b1e42-794b-4c71-93ac-5ed92488b67g' }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/9b1b1e42-794b-4c71-93ac-5ed92488b67f/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews" }]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' });

    await command.action(logger, { options: { groupName: groupName, role: "Owner" } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Owner"]
      }
    ]));
  });

  it('correctly lists all members in a Microsoft Entra group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "givenName": "Karl", "surname": "Matteson" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, role: "Member" } });

    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Member"]
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "givenName": "Karl",
        "surname": "Matteson",
        "roles": ["Member"]
      }
    ]));
  });

  it('correctly lists properties for all users in a Microsoft Entra group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Owners/microsoft.graph.user?$select=displayName,mail,id&$expand=memberof($select=id),memberof($select=displayName)`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Karl Matteson", "mail": "karl.matteson@contoso.onmicrosoft.com", "memberOf": [{ "displayName": "Life and Music", "id": "d6c88284-c598-468d-8074-56acaf3c0453" }] }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Members/microsoft.graph.user?$select=displayName,mail,id&$expand=memberof($select=id),memberof($select=displayName)`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Anne Matthews", "mail": "anne.matthews@contoso.onmicrosoft.com", "memberOf": [{ "displayName": "Life and Music", "id": "d6c88284-c598-468d-8074-56acaf3c0454" }] }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, properties: "displayName,mail,memberof/id,memberof/displayName" } });

    assert(loggerLogSpy.calledOnceWithExactly([
      { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Karl Matteson", "mail": "karl.matteson@contoso.onmicrosoft.com", "memberOf": [{ "displayName": "Life and Music", "id": "d6c88284-c598-468d-8074-56acaf3c0453" }], "roles": ["Owner"] },
      { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Anne Matthews", "mail": "anne.matthews@contoso.onmicrosoft.com", "memberOf": [{ "displayName": "Life and Music", "id": "d6c88284-c598-468d-8074-56acaf3c0454" }], "roles": ["Member"] }
    ]));
  });

  it('correctly lists all guest users in a Microsoft Entra group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname&$filter=userType%20eq%20'Guest'&$count=true`) {
        return {
          "value": []
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname&$filter=userType%20eq%20'Guest'&$count=true`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "annematthews_gmail.com#EXT#@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId, filter: "userType eq 'Guest'" } });

    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "annematthews_gmail.com#EXT#@contoso.onmicrosoft.com",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Member"]
      }
    ]));
  });

  it('throws an error when group by id cannot be found', async () => {
    const error = {
      error: {
        code: 'Request_ResourceNotFound',
        message: `Resource '${groupId}' does not exist or one of its queried reference-property objects are not present.`,
        innerError: {
          date: '2023-08-30T14:32:41',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, groupId: groupId } }),
      new CommandError(error.error.message));
  });
});