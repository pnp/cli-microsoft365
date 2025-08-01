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
import command from './serviceprincipal-grant-list.js';

describe(commands.SERVICEPRINCIPAL_GRANT_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const spoServicePrincipalDisplayName = 'SharePoint Online Web Client Extensibility';
  const spoServicePrincipalID = '00000000-0000-0000-0000-000000000000';
  const graphUrl = 'https://graph.microsoft.com/v1.0';
  const oauth2PermissionGrants = {
    value: [
      {
        clientId: '1e551032-3e2d-4d6b-9392-9b25451313a0',
        consentType: 'AllPrincipals',
        id: '50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw',
        principalId: null,
        resourceId: '1c444f1a-bba3-42f2-999f-4106c5b1c20c',
        scope: 'Group.ReadWrite.All'
      },
      {
        clientId: '1e551032-3e2d-4d6b-9392-9b25451313a0',
        consentType: 'AllPrincipals',
        id: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg',
        principalId: null,
        resourceId: 'dcf25ef3-e2df-4a77-839d-6b7857a11c78',
        scope: 'MyFiles.Read'
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    assert.strictEqual(command.name, commands.SERVICEPRINCIPAL_GRANT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists permissions granted to the service principal', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${graphUrl}/servicePrincipals?$filter=displayName eq '${spoServicePrincipalDisplayName}'&$select=id`) {
        return {
          value: [
            { id: spoServicePrincipalID }
          ]
        };
      }

      if (opts.url === `${graphUrl}/servicePrincipals/${spoServicePrincipalID}/oauth2PermissionGrants`) {
        return oauth2PermissionGrants;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(oauth2PermissionGrants.value));
  });

  it('returns error when the Service principal does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${graphUrl}/servicePrincipals?$filter=displayName eq '${spoServicePrincipalDisplayName}'&$select=id`) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Service principal '${spoServicePrincipalDisplayName}' not found`));
  });

  it('correctly handles random API error', async () => {
    const error = {
      "error": {
        "code": "Request_ResourceNotFound",
        "message": "Resource '20c5353f-acc6-424a-bf81-bc80fbd74cdb' does not exist or one of its queried reference-property objects are not present.",
        "innerError": {
          "date": "2025-07-07T13:43:51",
          "request-id": "5b16fbbb-61ef-432e-91d7-fa259efc184c",
          "client-request-id": "75b16fbbb-61ef-432e-91d7-fa259efc184c"
        }
      }
    };

    sinon.stub(request, 'get').rejects(error);
    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Resource '20c5353f-acc6-424a-bf81-bc80fbd74cdb' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });
});
