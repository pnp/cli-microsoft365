import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import command from './roledefinition-list.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';

describe(commands.ROLEDEFINITION_LIST, () => {
  const roleDefinitionsResponse = [
    {
      "id": "f28a1f50-f6e7-4571-818b-6a12f2af6b6c",
      "description": "Can manage all aspects of the SharePoint service.",
      "displayName": "SharePoint Administrator",
      "isBuiltIn": true,
      "isEnabled": true,
      "resourceScopes": [
        "/"
      ],
      "templateId": "f28a1f50-f6e7-4571-818b-6a12f2af6b6c",
      "version": "1",
      "rolePermissions": [
        {
          "allowedResourceActions": [
            "microsoft.azure.serviceHealth/allEntities/allTasks",
            "microsoft.azure.supportTickets/allEntities/allTasks",
            "microsoft.backup/oneDriveForBusinessProtectionPolicies/allProperties/allTasks",
            "microsoft.backup/oneDriveForBusinessRestoreSessions/allProperties/allTasks",
            "microsoft.backup/restorePoints/sites/allProperties/allTasks",
            "microsoft.backup/restorePoints/userDrives/allProperties/allTasks",
            "microsoft.backup/sharePointProtectionPolicies/allProperties/allTasks",
            "microsoft.backup/sharePointRestoreSessions/allProperties/allTasks",
            "microsoft.backup/siteProtectionUnits/allProperties/allTasks",
            "microsoft.backup/siteRestoreArtifacts/allProperties/allTasks",
            "microsoft.backup/userDriveProtectionUnits/allProperties/allTasks",
            "microsoft.backup/userDriveRestoreArtifacts/allProperties/allTasks",
            "microsoft.directory/groups/hiddenMembers/read",
            "microsoft.directory/groups.unified/basic/update",
            "microsoft.directory/groups.unified/create",
            "microsoft.directory/groups.unified/delete",
            "microsoft.directory/groups.unified/members/update",
            "microsoft.directory/groups.unified/owners/update",
            "microsoft.directory/groups.unified/restore",
            "microsoft.office365.migrations/allEntities/allProperties/allTasks",
            "microsoft.office365.network/performance/allProperties/read",
            "microsoft.office365.serviceHealth/allEntities/allTasks",
            "microsoft.office365.sharePoint/allEntities/allTasks",
            "microsoft.office365.supportTickets/allEntities/allTasks",
            "microsoft.office365.usageReports/allEntities/allProperties/read",
            "microsoft.office365.webPortal/allEntities/standard/read"
          ],
          "condition": null
        }
      ],
      "inheritsPermissionsFrom": [
        {
          "id": "88d8e3e3-8f55-4a1e-953a-9b9898b8876b"
        }
      ]
    },
    {
      "id": "abcd1234-de71-4623-b4af-96380a352509",
      "description": "Can read Bitlocker keys.",
      "displayName": "Bitlocker Keys Reader",
      "isBuiltIn": false,
      "isEnabled": true,
      "resourceScopes": [
        "/"
      ],
      "templateId": "abcd1234-de71-4623-b4af-96380a352509",
      "version": "1",
      "rolePermissions": [
        {
          "allowedResourceActions": [
            "microsoft.directory/bitlockerKeys/key/read"
          ],
          "condition": null
        }
      ],
      "inheritsPermissionsFrom": [
      ]
    }
  ];

  const roleDefinitionsLimitedResponse = [
    {
      "id": "f28a1f50-f6e7-4571-818b-6a12f2af6b6c",
      "displayName": "SharePoint Administrator",
      "isBuiltIn": true,
      "isEnabled": true
    },
    {
      "id": "abcd1234-de71-4623-b4af-96380a352509",
      "displayName": "Bitlocker Keys Reader",
      "isBuiltIn": false,
      "isEnabled": true
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
    assert.strictEqual(command.name, commands.ROLEDEFINITION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'isBuiltIn', 'isEnabled']);
  });

  it(`should get a list of Entra ID role definitions`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions`) {
        return {
          value: roleDefinitionsResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { verbose: true }
    });

    assert(loggerLogSpy.calledWith(roleDefinitionsResponse));
  });

  it(`should get a list of Entra ID role definitions with specified properties`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName,isBuiltIn,isEnabled`) {
        return {
          value: roleDefinitionsLimitedResponse
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { properties: 'id,displayName,isBuiltIn,isEnabled' }
    });

    assert(loggerLogSpy.calledWith(roleDefinitionsLimitedResponse));
  });

  it(`should get a list of filtered Entra ID role definitions`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=isBuiltIn eq false`) {
        return {
          value: [roleDefinitionsResponse[1]]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { filter: 'isBuiltIn eq false' }
    });

    assert(loggerLogSpy.calledWith([roleDefinitionsResponse[1]]));
  });

  it(`should get a list of filtered Entra ID role definitions with specified properties`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName,isBuiltIn,isEnabled&$filter=isBuiltIn eq false`) {
        return {
          value: [roleDefinitionsLimitedResponse[1]]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: { properties: 'id,displayName,isBuiltIn,isEnabled', filter: 'isBuiltIn eq false' }
    });

    assert(loggerLogSpy.calledWith([roleDefinitionsLimitedResponse[1]]));
  });

  it('handles error when retrieving role definitions failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions`) {
        throw { error: { message: 'An error has occurred' } };
      }
      throw `Invalid request`;
    });

    await assert.rejects(
      command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred')
    );
  });
});