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
import { aadGroup } from '../../../../utils/aadGroup.js';
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
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(true);
    auth.service.connected = true;
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
    auth.service.connected = false;
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

  it('correctly lists all users in a Microsoft 365 group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members?$select=id,displayName,userPrincipalName,userType`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000" } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Owner"
      },
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Member"
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "userType": "Member"
      }
    ]));
  });

  it('correctly lists all owners in a Microsoft 365 group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", role: "Owner" } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Owner"
      }
    ]));
  });

  it('correctly lists all members in a Microsoft 365 group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members?$select=id,displayName,userPrincipalName,userType`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000", role: "Member" } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Member"
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "userType": "Member"
      }
    ]));
  });

  it('correctly lists all users in a Microsoft 365 group (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners?$select=id,displayName,userPrincipalName,userType`) {
        return {
          "value": [{ "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" }]
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members?$select=id,displayName,userPrincipalName,userType`) {
        return {
          "value": [
            { "id": "00000000-0000-0000-0000-000000000000", "displayName": "Anne Matthews", "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com", "userType": "Member" },
            { "id": "00000000-0000-0000-0000-000000000001", "displayName": "Karl Matteson", "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com", "userType": "Member" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000" } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Owner"
      },
      {
        "id": "00000000-0000-0000-0000-000000000000",
        "displayName": "Anne Matthews",
        "userPrincipalName": "anne.matthews@contoso.onmicrosoft.com",
        "userType": "Member"
      },
      {
        "id": "00000000-0000-0000-0000-000000000001",
        "displayName": "Karl Matteson",
        "userPrincipalName": "karl.matteson@contoso.onmicrosoft.com",
        "userType": "Member"
      }
    ]));
  });

  it('correctly handles error when listing users', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000" } } as any),
      new CommandError('An error has occurred'));
  });

  it('throws error when the group is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(aadGroup.isUnifiedGroup);
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { groupId: groupId } } as any),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });
});
