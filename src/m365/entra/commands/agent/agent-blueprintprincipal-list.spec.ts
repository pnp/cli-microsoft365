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
import command, { options } from './agent-blueprintprincipal-list.js';

describe(commands.AGENT_BLUEPRINTPRINCIPAL_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  const response = [
    {
      "accountEnabled": true,
      "appDescription": null,
      "appDisplayName": "OcnAgentBluePrint02",
      "appId": "0e81af7f-b058-470a-84be-3a4f5a8014ca",
      "appOwnerOrganizationId": "70cf0eca-e1e8-4d78-9d03-62ecf811dfc2",
      "appRoleAssignmentRequired": false,
      "createdByAppId": "14d82eec-204b-4c2f-b7e8-296a70dab67e",
      "disabledByMicrosoftStatus": null,
      "displayName": "OcnAgentBluePrint02",
      "servicePrincipalNames": [
        "0e81af7f-b058-470a-84be-3a4f5a8014ca"
      ],
      "servicePrincipalType": "Application",
      "signInAudience": "AzureADMyOrg",
      "tags": [],
      "id": "59246ef4-b832-48eb-a5e0-87748132e12c",
      "appRoles": [],
      "info": {
        "logoUrl": null,
        "marketingUrl": null,
        "privacyStatementUrl": null,
        "supportUrl": null,
        "termsOfServiceUrl": null
      },
      "oauth2PermissionScopes": [],
      "verifiedPublisher": {
        "displayName": null,
        "verifiedPublisherId": null,
        "addedDateTime": null
      }
    }
  ];

  const limitedResponse = [
    {
      "id": "59246ef4-b832-48eb-a5e0-87748132e12c",
      "appDisplayName": "OcnAgentBluePrint02"
    }
  ];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    assert.strictEqual(command.name, commands.AGENT_BLUEPRINTPRINCIPAL_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'appDisplayName', 'appId']);
  });

  it(`should get a list of agent identity blueprints`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal`) {
        return {
          value: response
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });

    assert(
      loggerLogSpy.calledWith(response)
    );
  });

  it(`should get a list of administrative units with specified properties`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal?$select=id,appDisplayName`) {
        return {
          value: limitedResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ properties: 'id,appDisplayName' }) });

    assert(
      loggerLogSpy.calledWith(limitedResponse)
    );
  });

  it('handles error when retrieving administrative units list failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError('An error has occurred')
    );
  });
});
