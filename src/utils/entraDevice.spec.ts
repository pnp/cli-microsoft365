import assert from 'assert';
import sinon from 'sinon';
import { entraDevice } from './entraDevice.js';
import { cli } from "../cli/cli.js";
import request from "../request.js";
import { sinonUtil } from "./sinonUtil.js";
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';


describe('utils/entraDevice', () => {
  const deviceId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const secondDeviceId = 'fc33aa61-cf0e-1234-9506-f633347202ab';
  const displayName = 'AdeleVence-PC';
  const secondDisplayName = 'JohnDoe-PC';
  const invalidDisplayName = 'European';

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single device by name using getDeviceByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/devices?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            {
              id: deviceId,
              displayName: displayName
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraDevice.getDeviceByDisplayName(displayName);
    assert.deepStrictEqual(actual, { id: deviceId, displayName: displayName });
  });

  it('handles selecting single device when multiple devices with the specified name found using getDeviceByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/devices?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            { id: deviceId, displayName: displayName },
            { id: secondDeviceId, displayName: secondDisplayName }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: deviceId, displayName: displayName });

    const actual = await entraDevice.getDeviceByDisplayName(displayName);
    assert.deepStrictEqual(actual, { id: deviceId, displayName: displayName });
  });

  it('throws error message when no device was found using getDeviceByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/devices?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(entraDevice.getDeviceByDisplayName(invalidDisplayName)), Error(`The specified device '${invalidDisplayName}' does not exist.`);
  });

  it('throws error message when multiple devices were found using getDeviceByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/devices?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            { id: deviceId },
            { id: deviceId }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(entraDevice.getDeviceByDisplayName(displayName), Error(`Multiple devices with name '${displayName}' found. Found: ${deviceId}.`));
  });
});