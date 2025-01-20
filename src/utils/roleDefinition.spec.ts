import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { roleDefinition } from './roleDefinition.js';
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';

describe('utils/roleDefinition', () => {
  const id = '729827e3-9c14-49f7-bb1b-9608f156bbb8';
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
  const roleDefinitionLimitedResponse = {
    "id": "729827e3-9c14-49f7-bb1b-9608f156bbb8",
    "displayName": "Helpdesk Administrator"
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
  const exchangeRoleDefinitionResponse = {
    "id": "82fd214e-61ca-4dc7-98f6-090700bdb205",
    "description": "Allows the app to create, read, update, and delete email in all mailboxes without a signed-in user. Does not include permission to send mail",
    "displayName": "Application Mail.ReadWrite",
    "isEnabled": true,
    "version": "0.12 (14.0.451.0)",
    "isBuiltIn": true,
    "templateId": null,
    "allowedPrincipalTypes": "servicePrincipal",
    "rolePermissions": [
      {
        "allowedResourceActions": [
          "Mail.ReadWrite"
        ],
        "excludedResourceActions": [],
        "condition": null
      }
    ]
  };
  const secondExchangeRoleDefinitionResponse = {
    "id": "71ec103d-72db-5ed8-87e5-181611acc114",
    "description": "Allows the app to create, read, update, and delete events of all calendars without a signed-in user",
    "displayName": "Application Calendars.ReadWrite",
    "isEnabled": true,
    "version": "0.12 (14.0.451.0)",
    "isBuiltIn": true,
    "templateId": null,
    "allowedPrincipalTypes": "servicePrincipal",
    "rolePermissions": [
      {
        "allowedResourceActions": [
          "Calendars.ReadWrite"
        ],
        "excludedResourceActions": [],
        "condition": null
      }
    ]
  };

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single role definition by name using getDirectoryRoleDefinitionByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            roleDefinitionResponse
          ]
        };
      }

      throw 'Invalid Request';
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

  it('correctly get single role definition by name using getDirectoryRoleDefinitionByDisplayName with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id,displayName`) {
        return {
          value: [
            roleDefinitionLimitedResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await roleDefinition.getRoleDefinitionByDisplayName(displayName, 'id,displayName');
    assert.deepStrictEqual(actual, {
      "id": "729827e3-9c14-49f7-bb1b-9608f156bbb8",
      "displayName": "Helpdesk Administrator"
    });
  });

  it('handles selecting single role definition when multiple role definitions with the specified name found using getDirectoryRoleDefinitionByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            roleDefinitionResponse,
            customRoleDefinitionResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(roleDefinitionResponse);

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

  it('throws error message when no role definition was found using getDirectoryRoleDefinitionByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(roleDefinition.getRoleDefinitionByDisplayName(invalidDisplayName)), Error(`The specified role definition '${invalidDisplayName}' does not exist.`);
  });

  it('throws error message when multiple role definition were found using getDirectoryRoleDefinitionByDisplayName', async () => {
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

  it('correctly get single role definition by name using getRoleDefinitionById', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${id}`) {
        return roleDefinitionResponse;
      }

      throw 'Invalid Request';
    });

    const actual = await roleDefinition.getRoleDefinitionById(id);
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

  it('correctly get single role definition by name using getRoleDefinitionById with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/${id}?$select=id,displayName`) {
        return roleDefinitionResponse;
      }

      throw 'Invalid Request';
    });

    const actual = await roleDefinition.getRoleDefinitionById(id, 'id,displayName');
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

  it('correctly get single role definition by name using getExchangeRoleDefinitionByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            exchangeRoleDefinitionResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await roleDefinition.getExchangeRoleDefinitionByDisplayName(displayName);
    assert.deepStrictEqual(actual, exchangeRoleDefinitionResponse);
  });

  it('correctly get single role definition by name using getExchangeRoleDefinitionByDisplayName with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id,displayName`) {
        return {
          value: [
            exchangeRoleDefinitionResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await roleDefinition.getExchangeRoleDefinitionByDisplayName(displayName, 'id,displayName');
    assert.deepStrictEqual(actual, exchangeRoleDefinitionResponse);
  });

  it('handles selecting single role definition when multiple role definitions with the specified name found using getExchangeRoleDefinitionByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            exchangeRoleDefinitionResponse,
            secondExchangeRoleDefinitionResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(exchangeRoleDefinitionResponse);

    const actual = await roleDefinition.getExchangeRoleDefinitionByDisplayName(displayName);
    assert.deepStrictEqual(actual, exchangeRoleDefinitionResponse);
  });

  it('throws error message when no role definition was found using getExchangeRoleDefinitionByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(roleDefinition.getExchangeRoleDefinitionByDisplayName(invalidDisplayName)), Error(`The specified role definition '${invalidDisplayName}' does not exist.`);
  });

  it('throws error message when multiple role definition were found using getExchangeRoleDefinitionByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/roleDefinitions?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            exchangeRoleDefinitionResponse,
            secondExchangeRoleDefinitionResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(roleDefinition.getExchangeRoleDefinitionByDisplayName(displayName),
      Error(`Multiple role definitions with name '${displayName}' found. Found: ${exchangeRoleDefinitionResponse.id}, ${secondExchangeRoleDefinitionResponse.id}.`));
  });
});