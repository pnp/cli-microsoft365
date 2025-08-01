import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { vivaEngageRole } from './vivaEngageRole.js';
import { settingsNames } from '../settingsNames.js';


describe('utils/vivaEngage', () => {
  const roleId = 'ec759127-089f-4f91-8dfc-03a30b51cb38';
  const roleName = 'Network Admin';
  const invalidDisplayName = 'Network Admins';
  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single role id by name using getRoleIdByName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles`) {
        return {
          value: [
            {
              "id": "ec759127-089f-4f91-8dfc-03a30b51cb38",
              "displayName": "Network Admin"
            },
            {
              "id": "966b8ec4-6457-4f22-bd3c-5a2520e98f4a",
              "displayName": "Verified Admin"
            },
            {
              "id": "77aa47ad-96fe-4ecc-8024-fd1ac5e28f17",
              "displayName": "Corporate Communicator"
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await vivaEngageRole.getRoleIdByName(roleName);
    assert.deepStrictEqual(actual, roleId);
  });

  it('handles selecting single role when multiple roles with the specified name found using getRoleIdByName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles`) {
        return {
          value: [
            {
              "id": "ec759127-089f-4f91-8dfc-03a30b51cb38",
              "displayName": "Network Admin"
            },
            {
              "id": "966b8ec4-6457-4f22-bd3c-5a2520e98f4a",
              "displayName": "Network Admin"
            },
            {
              "id": "77aa47ad-96fe-4ecc-8024-fd1ac5e28f17",
              "displayName": "Corporate Communicator"
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: roleId });

    const actual = await vivaEngageRole.getRoleIdByName(roleName);
    assert.deepStrictEqual(actual, roleId);
  });

  it('throws error message when no role was found using getRoleIdByName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles`) {
        return {
          value: [
            {
              "id": "ec759127-089f-4f91-8dfc-03a30b51cb38",
              "displayName": "Network Admin"
            },
            {
              "id": "966b8ec4-6457-4f22-bd3c-5a2520e98f4a",
              "displayName": "Verified Admin"
            },
            {
              "id": "77aa47ad-96fe-4ecc-8024-fd1ac5e28f17",
              "displayName": "Corporate Communicator"
            }
          ] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(vivaEngageRole.getRoleIdByName(invalidDisplayName)), Error(`The specified Viva Engage role '${invalidDisplayName}' does not exist.`);
  });

  it('throws error message when multiple communities were found using getRoleIdByName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/employeeExperience/roles`) {
        return {
          value: [
            {
              "id": "ec759127-089f-4f91-8dfc-03a30b51cb38",
              "displayName": "Network Admin"
            },
            {
              "id": "966b8ec4-6457-4f22-bd3c-5a2520e98f4a",
              "displayName": "Network Admin"
            },
            {
              "id": "77aa47ad-96fe-4ecc-8024-fd1ac5e28f17",
              "displayName": "Corporate Communicator"
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(vivaEngageRole.getRoleIdByName(roleName),
      Error(`Multiple Viva Engage roles with name '${roleName}' found. Found: ${roleId}, 966b8ec4-6457-4f22-bd3c-5a2520e98f4a.`));
  });
});