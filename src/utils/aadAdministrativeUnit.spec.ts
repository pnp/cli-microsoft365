import assert from 'assert';
import sinon from 'sinon';
import { aadAdministrativeUnit } from './aadAdministrativeUnit.js';
import { cli } from "../cli/cli.js";
import request from "../request.js";
import { sinonUtil } from "./sinonUtil.js";
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';
import { aadUser } from './aadUser.js';
import { aadGroup } from './aadGroup.js';
import { aadDevice } from './aadDevice.js';


describe('utils/aadAdministrativeUnit', () => {
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const secondAdministrativeUnitId = 'fc33aa61-cf0e-1234-9506-f633347202ab';
  const displayName = 'European Division';
  const secondDisplayName = 'Asian Division';
  const invalidDisplayName = 'European';
  const userId = '64131a70-beb9-4ccb-b590-4401e58446ec';
  const userName = 'PradeepG@4wrvkx.onmicrosoft.com';
  const groupId = 'c121c70b-deb1-43f7-8298-9111bf3036b4';
  const groupName = 'Mark 8 Project Team';
  const deviceId = '3f9fd7c3-73ad-4ce3-b053-76bb8252964d';
  const deviceName = 'AdeleVence-PC';

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      aadUser.getUserIdByUpn,
      aadGroup.getGroupIdByDisplayName,
      aadDevice.getDeviceByDisplayName
    ]);
  });

  it('correctly get single administrative unit id by name using getAdministrativeUnitByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            {
              id: administrativeUnitId,
              displayName: displayName
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await aadAdministrativeUnit.getAdministrativeUnitByDisplayName(displayName);
    assert.deepStrictEqual(actual, { id: administrativeUnitId, displayName: displayName });
  });

  it('handles selecting single administrative unit when multiple administrative units with the specified name found using getAdministrativeUnitByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            { id: administrativeUnitId, displayName: displayName },
            { id: secondAdministrativeUnitId, displayName: secondDisplayName }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: administrativeUnitId, displayName: displayName });

    const actual = await aadAdministrativeUnit.getAdministrativeUnitByDisplayName(displayName);
    assert.deepStrictEqual(actual, { id: administrativeUnitId, displayName: displayName });
  });

  it('throws error message when no administrative unit was found using getAdministrativeUnitByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(aadAdministrativeUnit.getAdministrativeUnitByDisplayName(invalidDisplayName)), Error(`The specified administrative unit '${invalidDisplayName}' does not exist.`);
  });

  it('throws error message when multiple administrative units were found using getAdministrativeUnitByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            { id: administrativeUnitId },
            { id: administrativeUnitId }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(aadAdministrativeUnit.getAdministrativeUnitByDisplayName(displayName), Error(`Multiple administrative units with name '${displayName}' found. Found: ${administrativeUnitId}.`));
  });

  it('throws error when no member was found using getMemberIdByName', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').throwsException(Error(`The specified user with user name m365 does not exist.`));
    sinon.stub(aadGroup, 'getGroupIdByDisplayName').throwsException(Error(`The specified group 'm365' does not exist.`));
    sinon.stub(aadDevice, 'getDeviceByDisplayName').throwsException(Error(`The specified device 'm365' does not exist.`));

    await assert.rejects(aadAdministrativeUnit.getMemberIdByName('m365'), Error(`The specified member 'm365' does not exist.`));
  });

  it('get member id using getMemberIdByName when member specified by user name', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').withArgs(userName).resolves(userId);
    const aadGroupStub = sinon.stub(aadGroup, 'getGroupIdByDisplayName');
    const aadDeviceStub = sinon.stub(aadDevice, 'getDeviceByDisplayName');

    const actual = await aadAdministrativeUnit.getMemberIdByName(userName);

    assert.strictEqual(actual, userId);
    assert.strictEqual(aadGroupStub.notCalled, true);
    assert.strictEqual(aadDeviceStub.notCalled, true);
  });

  it('get member id using getMemberIdByName when member specified by group name', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').withArgs(groupName).throwsException(Error(`The specified user with user name ${groupName} does not exist.`));
    sinon.stub(aadGroup, 'getGroupIdByDisplayName').withArgs(groupName).resolves(groupId);
    sinon.stub(aadDevice, 'getDeviceByDisplayName').withArgs(groupName).throwsException(Error(`The specified device '${groupName}' does not exist.`));

    const actual = await aadAdministrativeUnit.getMemberIdByName(groupName);

    assert.strictEqual(actual, groupId);
  });

  it('get member id using getMemberIdByName when member specified by device name', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').withArgs(deviceName).throwsException(Error(`The specified user with user name ${deviceName} does not exist.`));
    sinon.stub(aadGroup, 'getGroupIdByDisplayName').withArgs(deviceName).throwsException(Error(`The specified group '${deviceName}' does not exist.`));
    sinon.stub(aadDevice, 'getDeviceByDisplayName').withArgs(deviceName).resolves({ id: deviceId, displayName: deviceName });

    const actual = await aadAdministrativeUnit.getMemberIdByName(deviceName);

    assert.strictEqual(actual, deviceId);
  });

  it('handle selecting single member using getMemberIdByName when group and device with the same name was found', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').withArgs(deviceName).throwsException(Error(`The specified user with user name ${deviceName} does not exist.`));
    sinon.stub(aadGroup, 'getGroupIdByDisplayName').withArgs(deviceName).resolves(groupId);
    sinon.stub(aadDevice, 'getDeviceByDisplayName').withArgs(deviceName).resolves({ id: deviceId, displayName: deviceName });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: deviceId, displayName: deviceName });

    const actual = await aadAdministrativeUnit.getMemberIdByName(deviceName);

    assert.strictEqual(actual, deviceId);
  });
});