import assert from 'assert';
import sinon from 'sinon';
import { entraAdministrativeUnit } from './entraAdministrativeUnit.js';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';


describe('utils/entraAdministrativeUnit', () => {
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const secondAdministrativeUnitId = 'fc33aa61-cf0e-1234-9506-f633347202ab';
  const displayName = 'European Division';
  const secondDisplayName = 'Asian Division';
  const invalidDisplayName = 'European';

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
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

    const actual = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(displayName);
    assert.deepStrictEqual(actual, { id: administrativeUnitId, displayName: displayName });
  });

  it('correctly get single administrative unit id by name using getAdministrativeUnitByDisplayName with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'?$select=id,displayName`) {
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

    const actual = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(displayName, 'id,displayName');
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

    const actual = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(displayName);
    assert.deepStrictEqual(actual, { id: administrativeUnitId, displayName: displayName });
  });

  it('throws error message when no administrative unit was found using getAdministrativeUnitByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(entraAdministrativeUnit.getAdministrativeUnitByDisplayName(invalidDisplayName)), Error(`The specified administrative unit '${invalidDisplayName}' does not exist.`);
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

    await assert.rejects(entraAdministrativeUnit.getAdministrativeUnitByDisplayName(displayName), Error(`Multiple administrative units with name '${displayName}' found. Found: ${administrativeUnitId}.`));
  });
});