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
import command from './m365group-user-list.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import aadCommands from '../../aadCommands.js';

describe(commands.M365GROUP_USER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(true);
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_USER_LIST);
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
    assert.deepStrictEqual(alias, [aadCommands.M365GROUP_USER_LIST]);
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
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid groupId and Owner role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        role: 'Owner'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid groupId and Member role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        role: 'Member'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid groupId and Guest role specified', async () => {
    const actual = await command.validate({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        role: 'Guest'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly lists all users in a Microsoft 365 group by group id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "givenName": "Karl", "surname": "Matteson", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, groupId: "00000000-0000-0000-0000-000000000000" } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Owner",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Owner", "Member"]
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "userType": "Member",
        "givenName": "Karl",
        "surname": "Matteson",
        "roles": ["Member"]
      }
    ]));
  });

  it('correctly lists all users in a Microsoft 365 group by group name', async () => {
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves('00000000-0000-0000-0000-000000000000');

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "givenName": "Karl", "surname": "Matteson", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, groupDisplayName: "CLI Test Group" } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Owner",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Owner", "Member"]
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "userType": "Member",
        "givenName": "Karl",
        "surname": "Matteson",
        "roles": ["Member"]
      }
    ]));
  });

  it('correctly lists all owners in a Microsoft 365 group by group id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", role: "Owner" } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Owner",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Owner"]
      }
    ]));
  });

  it('correctly lists all members in a Microsoft 365 group by group id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "givenName": "Karl", "surname": "Matteson", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", role: "Member" } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Member",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Member"]
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "userType": "Member",
        "givenName": "Karl",
        "surname": "Matteson",
        "roles": ["Member"]
      }
    ]));
  });

  it('correctly lists all guests in a Microsoft 365 group by group id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "givenName": "Karl", "surname": "Matteson", "userType": "Member" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "annematthews_gmail.com#EXT#@nachan365.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Guest" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "givenName": "Karl", "surname": "Matteson", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", role: "Guest" } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "annematthews_gmail.com#EXT#@nachan365.onmicrosoft.com",
        "userType": "Guest",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Member"]
      }
    ]));
  });

  it('correctly lists all users in a Microsoft 365 group by group id (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Member" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "givenName": "Karl", "surname": "Matteson", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000" } });
    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Owner",
        "givenName": "Anne",
        "surname": "Matthews",
        "roles": ["Owner", "Member"]
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "userType": "Member",
        "givenName": "Karl",
        "surname": "Matteson",
        "roles": ["Member"]
      }
    ]));
  });

  it('correctly lists properties for all users in a Microsoft 365 group', async () => {
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

    await command.action(logger, { options: { groupId: "2c1ba4c4-cd9b-4417-832f-92a34bc34b2a", properties: "displayName,mail,memberof/id,memberof/displayName" } });

    assert(loggerLogSpy.calledOnceWithExactly([
      { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Karl Matteson", "mail": "karl.matteson@contoso.onmicrosoft.com", "memberOf": [{ "displayName": "Life and Music", "id": "d6c88284-c598-468d-8074-56acaf3c0453" }], "roles": ["Owner"], "userType": "Owner" },
      { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Anne Matthews", "mail": "anne.matthews@contoso.onmicrosoft.com", "memberOf": [{ "displayName": "Life and Music", "id": "d6c88284-c598-468d-8074-56acaf3c0454" }], "roles": ["Member"] }
    ]));
  });

  it('correctly lists all guest users in a Microsoft 365 group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Owners/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType&$filter=userType%20eq%20'Guest'&$count=true`) {
        return {
          "value": []
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/2c1ba4c4-cd9b-4417-832f-92a34bc34b2a/Members/microsoft.graph.user?$select=id,displayName,userPrincipalName,givenName,surname,userType&$filter=userType%20eq%20'Guest'&$count=true`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "annematthews_gmail.com#EXT#@contoso.onmicrosoft.com", "givenName": "Anne", "surname": "Matthews", "userType": "Guest" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: "2c1ba4c4-cd9b-4417-832f-92a34bc34b2a", filter: "userType eq 'Guest'" } });

    assert(loggerLogSpy.calledOnceWithExactly([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "annematthews_gmail.com#EXT#@contoso.onmicrosoft.com",
        "givenName": "Anne",
        "surname": "Matthews",
        "userType": "Guest",
        "roles": ["Member"]
      }
    ]));
  });

  it('correctly handles error when listing users', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000" } } as any),
      new CommandError('An error has occurred'));
  });

  it('throws error when the group by id is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { verbose: true, groupId: groupId } } as any),
      new CommandError(`Specified group '${groupId}' is not a Microsoft 365 group.`));
  });

  it('throws error when the group by name is not a unified group', async () => {
    const groupDisplayName = 'CLI Test Group';

    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { verbose: true, groupDisplayName: groupDisplayName } } as any),
      new CommandError(`Specified group '${groupDisplayName}' is not a Microsoft 365 group.`));
  });
});
