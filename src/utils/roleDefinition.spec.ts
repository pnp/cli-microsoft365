import assert from 'assert';
import sinon from 'sinon';
import { Cli } from "../cli/Cli.js";
import request from "../request.js";
import { sinonUtil } from "./sinonUtil.js";
import { roleDefinition } from './roleDefinition.js';
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';

describe('utils/roleDefinition', () => {
  const displayName = 'Helpdesk Administrator';
  const invalidDisplayName = 'Helpdeks Administratr';
  const roleDefinitionResponse = {
    "id": "729827e3-9c14-49f7-bb1b-9608f156bbb8",
    "description": "Can reset passwords for non-administrators and Helpdesk Administrators.",
    "displayName": "Helpdesk Administrator",
    "isBuiltIn": true,
    "isEnabled": true,
    "templateId": "729827e3-9c14-49f7-bb1b-9608f156bbb8",
    "version": "1",
    "rolePermissions": [
      {
        "allowedResourceActions": [
          "microsoft.directory/users/invalidateAllRefreshTokens",
          "microsoft.directory/users/bitLockerRecoveryKeys/read",
          "microsoft.directory/users/password/update",
          "microsoft.azure.serviceHealth/allEntities/allTasks",
          "microsoft.azure.supportTickets/allEntities/allTasks",
          "microsoft.office365.webPortal/allEntities/standard/read",
          "microsoft.office365.serviceHealth/allEntities/allTasks",
          "microsoft.office365.supportTickets/allEntities/allTasks"
        ],
        "condition": null
      }
    ],
    "inheritsPermissionsFrom": [
      {
        "id": "88d8e3e3-8f55-4a1e-953a-9b9898b8876b"
      }
    ]
  };
  const customRoleDefinitionResponse = {
    "id": "129827e3-9c14-49f7-bb1b-9608f156bbb8",
    "description": "Can update passwords for non-administrators and Helpdesk Administrators.",
    "displayName": "Helpdesk Administrator",
    "isBuiltIn": false,
    "isEnabled": true,
    "templateId": "729827e3-9c14-49f7-bb1b-9608f156bbb8",
    "version": "1",
    "rolePermissions": [
      {
        "allowedResourceActions": [
          "microsoft.directory/users/invalidateAllRefreshTokens",
          "microsoft.directory/users/bitLockerRecoveryKeys/read",
          "microsoft.directory/users/password/update"
        ],
        "condition": null
      }
    ],
    "inheritsPermissionsFrom": [
      {
        "id": "88d8e3e3-8f55-4a1e-953a-9b9898b8876b"
      }
    ]
  };
  let cli: Cli;

  before(() => {
    cli = Cli.getInstance();
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single role definition by name using getRoleDefinitionByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            roleDefinitionResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await roleDefinition.getRoleDefinitionByDisplayName(displayName);
    assert.deepStrictEqual(actual, {
      "id": "729827e3-9c14-49f7-bb1b-9608f156bbb8",
      "description": "Can reset passwords for non-administrators and Helpdesk Administrators.",
      "displayName": "Helpdesk Administrator",
      "isBuiltIn": true,
      "isEnabled": true,
      "templateId": "729827e3-9c14-49f7-bb1b-9608f156bbb8",
      "version": "1",
      "rolePermissions": [
        {
          "allowedResourceActions": [
            "microsoft.directory/users/invalidateAllRefreshTokens",
            "microsoft.directory/users/bitLockerRecoveryKeys/read",
            "microsoft.directory/users/password/update",
            "microsoft.azure.serviceHealth/allEntities/allTasks",
            "microsoft.azure.supportTickets/allEntities/allTasks",
            "microsoft.office365.webPortal/allEntities/standard/read",
            "microsoft.office365.serviceHealth/allEntities/allTasks",
            "microsoft.office365.supportTickets/allEntities/allTasks"
          ],
          "condition": null
        }
      ],
      "inheritsPermissionsFrom": [
        {
          "id": "88d8e3e3-8f55-4a1e-953a-9b9898b8876b"
        }
      ]
    });
  });

  it('handles selecting single role definition when multiple role definitions with the specified name found using getRoleDefinitionByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            roleDefinitionResponse,
            customRoleDefinitionResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves(roleDefinitionResponse);

    const actual = await roleDefinition.getRoleDefinitionByDisplayName(displayName);
    assert.deepStrictEqual(actual, {
      "id": "729827e3-9c14-49f7-bb1b-9608f156bbb8",
      "description": "Can reset passwords for non-administrators and Helpdesk Administrators.",
      "displayName": "Helpdesk Administrator",
      "isBuiltIn": true,
      "isEnabled": true,
      "templateId": "729827e3-9c14-49f7-bb1b-9608f156bbb8",
      "version": "1",
      "rolePermissions": [
        {
          "allowedResourceActions": [
            "microsoft.directory/users/invalidateAllRefreshTokens",
            "microsoft.directory/users/bitLockerRecoveryKeys/read",
            "microsoft.directory/users/password/update",
            "microsoft.azure.serviceHealth/allEntities/allTasks",
            "microsoft.azure.supportTickets/allEntities/allTasks",
            "microsoft.office365.webPortal/allEntities/standard/read",
            "microsoft.office365.serviceHealth/allEntities/allTasks",
            "microsoft.office365.supportTickets/allEntities/allTasks"
          ],
          "condition": null
        }
      ],
      "inheritsPermissionsFrom": [
        {
          "id": "88d8e3e3-8f55-4a1e-953a-9b9898b8876b"
        }
      ]
    });
  });

  it('throws error message when no role definition was found using getRoleDefinitionByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(roleDefinition.getRoleDefinitionByDisplayName(invalidDisplayName)), Error(`The specified role definition '${invalidDisplayName}' does not exist.`);
  });

  it('throws error message when multiple role definition were found using getRoleDefinitionByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            roleDefinitionResponse,
            customRoleDefinitionResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(roleDefinition.getRoleDefinitionByDisplayName(displayName),
      Error(`Multiple role definitions with name '${displayName}' found. Found: ${roleDefinitionResponse.id}, ${customRoleDefinitionResponse.id}.`));
  });
});