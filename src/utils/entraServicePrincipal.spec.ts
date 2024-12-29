import assert from 'assert';
import sinon from 'sinon';
import { entraServicePrincipal } from './entraServicePrincipal.js';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';

describe('utils/entraServicePrincipal', () => {
  const servicePrincipalId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const appId = '7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091';
  const appName = 'ContosoApp';
  const secondServicePrincipalId = 'fc33aa61-cf0e-1234-9506-f633347202ac';
  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single service principal id by appId using getServicePrincipalIdFromFromAppId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=id`) {
        return {
          value: [
            {
              id: servicePrincipalId
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraServicePrincipal.getServicePrincipalIdFromAppId(appId);
    assert.deepStrictEqual(actual, servicePrincipalId);
  });

  it('correctly get single service principal id by appId using getServicePrincipalIdFromAppName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'&$select=id`) {
        return {
          value: [
            {
              id: servicePrincipalId
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraServicePrincipal.getServicePrincipalIdFromAppName(appName);
    assert.deepStrictEqual(actual, servicePrincipalId);
  });

  it('handles selecting single service principal when multiple servicePrincipals with the specified name found using getServicePrincipalIdFromAppName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'&$select=id`) {
        return {
          value: [
            { id: servicePrincipalId },
            { id: secondServicePrincipalId }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: secondServicePrincipalId });

    const actual = await entraServicePrincipal.getServicePrincipalIdFromAppName(appName);
    assert.deepStrictEqual(actual, secondServicePrincipalId);
  });

  it('throws error message when no service principal was found using getServicePrincipalIdFromFromAppId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'&$select=id`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(entraServicePrincipal.getServicePrincipalIdFromAppId(appId)), Error(`App with appId '${appId}' not found in Microsoft Entra ID`);
  });

  it('throws error message when no service principal was found using getServicePrincipalIdFromAppName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'&$select=id`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(entraServicePrincipal.getServicePrincipalIdFromAppName(appName)), Error(`Service principal with name '${appName}' not found in Microsoft Entra ID`);
  });

  it('throws error message when multiple service principals were found using getServicePrincipalIdFromAppName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'&$select=id`) {
        return {
          value: [
            { id: servicePrincipalId },
            { id: secondServicePrincipalId }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(entraServicePrincipal.getServicePrincipalIdFromAppName(appName), Error(`Multiple service principals with name '${appName}' found in Microsoft Entra ID. Found: ${servicePrincipalId}, ${secondServicePrincipalId}.`));
  });
});