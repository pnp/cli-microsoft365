import assert from 'assert';
import sinon from 'sinon';
import { aadAdministrativeUnit } from './aadAdministrativeUnit.js';
import { Cli } from "../cli/Cli.js";
import request from "../request.js";
import { sinonUtil } from "./sinonUtil.js";
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';


describe('utils/aadAdministrativeUnit', () => {
  const administrativeUnitId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const secondAdministrativeUnitId = 'fc33aa61-cf0e-1234-9506-f633347202ab';
  const displayName = 'European Division';
  const invalidDisplayName = 'European';

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

  it('correctly get single administrative unit id by name using getAdministrativeUnitIdByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
        return {
          value: [
            { id: administrativeUnitId }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await aadAdministrativeUnit.getAdministrativeUnitIdByDisplayName(displayName);
    assert.deepStrictEqual(actual, administrativeUnitId);
  });

  it('handles selecting single administrative unit when multiple administrative units with the specified name found using getAdministrativeUnitIdByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
        return {
          value: [
            { id: administrativeUnitId },
            { id: secondAdministrativeUnitId }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves({ id: administrativeUnitId });

    const actual = await aadAdministrativeUnit.getAdministrativeUnitIdByDisplayName(displayName);
    assert.deepStrictEqual(actual, administrativeUnitId);
  });

  it('throws error message when no administrative unit was found using getAdministrativeUnitIdByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'&$select=id`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await assert.rejects(aadAdministrativeUnit.getAdministrativeUnitIdByDisplayName(invalidDisplayName)), Error(`The specified administrative unit '${invalidDisplayName}' does not exist.`);
  });

  it('throws error message when multiple administrative units were found using getAdministrativeUnitIdByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`) {
        return {
          value: [
            { id: administrativeUnitId },
            { id: administrativeUnitId }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(aadAdministrativeUnit.getAdministrativeUnitIdByDisplayName(displayName), Error(`Multiple administrative units with name '${displayName}' found. Found: ${administrativeUnitId}.`));
  });
});