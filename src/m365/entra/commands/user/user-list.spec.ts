import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-list.js';
import aadCommands from '../../aadCommands.js';

describe(commands.USER_LIST, () => {
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
    assert.strictEqual(command.name, commands.USER_LIST);
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
    assert.deepStrictEqual(alias, [aadCommands.USER_LIST]);
  });

  it('lists users in the tenant with the default properties (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$top=100`) {
        return {
          "value": [
            { "id": "1f5595b2-aa07-445d-9801-a45ea18160b2", "displayName": "Aarif Sherzai", "mail": "AarifS@contoso.onmicrosoft.com", "userPrincipalName": "AarifS@contoso.onmicrosoft.com" },
            { "id": "717f1683-00fa-488c-b68d-5d0051f6bcfa", "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "userPrincipalName": "AchimM@contoso.onmicrosoft.com" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([
      { "id": "1f5595b2-aa07-445d-9801-a45ea18160b2", "displayName": "Aarif Sherzai", "mail": "AarifS@contoso.onmicrosoft.com", "userPrincipalName": "AarifS@contoso.onmicrosoft.com" },
      { "id": "717f1683-00fa-488c-b68d-5d0051f6bcfa", "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "userPrincipalName": "AchimM@contoso.onmicrosoft.com" }
    ]));
  });

  it('retrieves only the specified user properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=displayName,mail,id&$top=100`) {
        return {
          "value": [
            { "displayName": "Aarif Sherzai", "mail": "AarifS@contoso.onmicrosoft.com", "id": "1f5595b2-aa07-445d-9801-a45ea18160b2" }, { "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "id": "717f1683-00fa-488c-b68d-5d0051f6bcfa" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { properties: 'displayName,mail' } });
    assert(loggerLogSpy.calledWith([
      { "displayName": "Aarif Sherzai", "mail": "AarifS@contoso.onmicrosoft.com", "id": "1f5595b2-aa07-445d-9801-a45ea18160b2" }, { "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "id": "717f1683-00fa-488c-b68d-5d0051f6bcfa" }
    ]));
  });

  it('retrieves properties for all users with properties option includes values with a slash', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=displayName,id&$expand=manager($select=displayName),manager($select=department)&$top=100`) {
        return {
          "value": [
            { "displayName": "Aarif Sherzai", "manager": { "displayName": "Jon Doe", "department": "IT" } }, { "displayName": "Achim Maier", "manager": { "displayName": "Jon Doe", "department": "IT" } }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { properties: 'displayName,manager/displayName,manager/department' } });
    assert(loggerLogSpy.calledWith([
      { "displayName": "Aarif Sherzai", "manager": { "displayName": "Jon Doe", "department": "IT" } }, { "displayName": "Achim Maier", "manager": { "displayName": "Jon Doe", "department": "IT" } }
    ]));
  });

  it('filters users by one property', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$filter=startsWith(surname, 'M')&$top=100`) {
        return {
          "value": [
            { "id": "1f5595b2-aa07-445d-9801-a45ea18160b2", "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "userPrincipalName": "AchimM@contoso.onmicrosoft.com" }, { "id": "0fe76bf5-222b-48f8-a5c1-a3a03b96d472", "displayName": "Karl Matteson", "mail": "KarlM@contoso.onmicrosoft.com", "userPrincipalName": "KarlM@contoso.onmicrosoft.com" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { surname: 'M' } });
    assert(loggerLogSpy.calledWith([
      { "id": "1f5595b2-aa07-445d-9801-a45ea18160b2", "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "userPrincipalName": "AchimM@contoso.onmicrosoft.com" }, { "id": "0fe76bf5-222b-48f8-a5c1-a3a03b96d472", "displayName": "Karl Matteson", "mail": "KarlM@contoso.onmicrosoft.com", "userPrincipalName": "KarlM@contoso.onmicrosoft.com" }
    ]));
  });

  it('filters users by multiple properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$filter=startsWith(surname, 'M') and startsWith(givenName, 'A')&$top=100`) {
        return {
          "value": [
            { "id": "1f5595b2-aa07-445d-9801-a45ea18160b2", "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "userPrincipalName": "AchimM@contoso.onmicrosoft.com" }, { "id": "7f50c7d9-916b-4da9-949e-09a431de2646", "displayName": "Anne Matthews", "mail": "AnneM@contoso.onmicrosoft.com", "userPrincipalName": "AnneM@contoso.onmicrosoft.com" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { surname: 'M', givenName: 'A' } });
    assert(loggerLogSpy.calledWith([
      { "id": "1f5595b2-aa07-445d-9801-a45ea18160b2", "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "userPrincipalName": "AchimM@contoso.onmicrosoft.com" }, { "id": "7f50c7d9-916b-4da9-949e-09a431de2646", "displayName": "Anne Matthews", "mail": "AnneM@contoso.onmicrosoft.com", "userPrincipalName": "AnneM@contoso.onmicrosoft.com" }
    ]));
  });

  it('lists users in the tenant with the guest type', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$filter=startsWith(surname, 'S') and userType eq 'Guest'&$top=100`) {
        return {
          "value": [
            { "id": "7dc52cef-c513-4a53-bd43-93e9f6727911", "displayName": "Aarif Sherzai", "mail": "AarifS@fabrikam.onmicrosoft.com", "userPrincipalName": "AarifS_fabrikam.onmicrosoft.com#EXT#@contoso.onmicrosoft.com" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { surname: 'S', type: 'Guest' } });
    assert(loggerLogSpy.calledWith([
      { "id": "7dc52cef-c513-4a53-bd43-93e9f6727911", "displayName": "Aarif Sherzai", "mail": "AarifS@fabrikam.onmicrosoft.com", "userPrincipalName": "AarifS_fabrikam.onmicrosoft.com#EXT#@contoso.onmicrosoft.com" }
    ]));
  });

  it('lists users in the tenant with the guest type and surname', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$filter=userType eq 'Guest'&$top=100`) {
        return {
          "value": [
            { "id": "7dc52cef-c513-4a53-bd43-93e9f6727911", "displayName": "Aarif Sherzai", "mail": "AarifS@fabrikam.onmicrosoft.com", "userPrincipalName": "AarifS_fabrikam.onmicrosoft.com#EXT#@contoso.onmicrosoft.com" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { type: 'Guest' } });
    assert(loggerLogSpy.calledWith([
      { "id": "7dc52cef-c513-4a53-bd43-93e9f6727911", "displayName": "Aarif Sherzai", "mail": "AarifS@fabrikam.onmicrosoft.com", "userPrincipalName": "AarifS_fabrikam.onmicrosoft.com#EXT#@contoso.onmicrosoft.com" }
    ]));
  });

  it('escapes special characters in filters', async () => {
    const displayName = 'O\'Brien';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$filter=startsWith(displayName, '${formatting.encodeQueryParameter(displayName)}')&$top=100`) {
        return {
          "value": []
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: displayName } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('ignores global options in filters', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$filter=startsWith(surname, 'M') and startsWith(givenName, 'A')&$top=100`) {
        return {
          "value": [
            { "id": "1f5595b2-aa07-445d-9801-a45ea18160b2", "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "userPrincipalName": "AchimM@contoso.onmicrosoft.com" }, { "id": "7f50c7d9-916b-4da9-949e-09a431de2646", "displayName": "Anne Matthews", "mail": "AnneM@contoso.onmicrosoft.com", "userPrincipalName": "AnneM@contoso.onmicrosoft.com" }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        surname: 'M',
        givenName: 'A',
        output: "json",
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith([
      { "id": "1f5595b2-aa07-445d-9801-a45ea18160b2", "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "userPrincipalName": "AchimM@contoso.onmicrosoft.com" }, { "id": "7f50c7d9-916b-4da9-949e-09a431de2646", "displayName": "Anne Matthews", "mail": "AnneM@contoso.onmicrosoft.com", "userPrincipalName": "AnneM@contoso.onmicrosoft.com" }
    ]));
  });

  it('handles error when retrieving second page of users', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$top=100`) {
        return {
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$top=100&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27",
          "value": [
            { "id": "1f5595b2-aa07-445d-9801-a45ea18160b2", "displayName": "Achim Maier", "mail": "AchimM@contoso.onmicrosoft.com", "userPrincipalName": "AchimM@contoso.onmicrosoft.com" }, { "id": "7f50c7d9-916b-4da9-949e-09a431de2646", "displayName": "Anne Matthews", "mail": "AnneM@contoso.onmicrosoft.com", "userPrincipalName": "AnneM@contoso.onmicrosoft.com" }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName&$top=100&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27`) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('handles error when option to filter by specified without a value (flag)', async () => {
    await assert.rejects(command.action(logger, { options: { surname: true } } as any), new CommandError('Specify value for the surname property'));
  });

  it('allows unknown properties', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });
});
