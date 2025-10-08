import assert from 'assert';
import sinon from 'sinon';
import { entraApp } from './entraApp.js';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';

describe('utils/entraApp', () => {
  const appObjectId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const appId = '7f5df2f4-9ed6-4df7-86d7-eefbfc4ab091';
  const appName = 'ContosoApp';
  const secondAppObjectId = 'fc33aa61-cf0e-1234-9506-f633347202ac';

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single app object id by appId using getAppRegistrationByAppId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'`) {
        return {
          value: [
            {
              id: appObjectId
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraApp.getAppRegistrationByAppId(appId);
    assert.deepStrictEqual(actual.id, appObjectId);
  });

  it('correctly get single app object id by appId using getAppRegistrationByAppName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'`) {
        return {
          value: [
            {
              id: appObjectId
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraApp.getAppRegistrationByAppName(appName);
    assert.deepStrictEqual(actual.id, appObjectId);
  });

  it('correctly get single app by appObjectId using getAppRegistrationByObjectId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}`) {
        return {
          id: appObjectId,
          displayName: appName
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraApp.getAppRegistrationByObjectId(appObjectId);
    assert.deepStrictEqual(actual.displayName, appName);
  });

  it('correctly get single app with specified properties by appObjectId using getAppRegistrationByObjectId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}?$select=id,displayName`) {
        return {
          id: appObjectId,
          displayName: appName
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraApp.getAppRegistrationByObjectId(appObjectId, ["id", "displayName"]);
    assert.deepStrictEqual(actual.displayName, appName);
  });

  it('handles selecting single application when multiple applications with the specified name found using getAppRegistrationByAppName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'&$select=id`) {
        return {
          value: [
            { id: appObjectId },
            { id: secondAppObjectId }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: secondAppObjectId });

    const actual = await entraApp.getAppRegistrationByAppName(appName, ['id']);
    assert.deepStrictEqual(actual.id, secondAppObjectId);
  });

  it('throws error message when no application was found using getAppRegistrationByAppId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications?$filter=appId eq '${appId}'&$select=id,displayName`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(entraApp.getAppRegistrationByAppId(appId, ['id', 'displayName']),
      new Error(`App with appId '${appId}' not found in Microsoft Entra ID.`));
  });

  it('throws error message when no application was found using getAppRegistrationByAppName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'&$select=id`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(entraApp.getAppRegistrationByAppName(appName, ['id']),
      new Error(`App with name '${appName}' not found in Microsoft Entra ID.`));
  });

  it('throws error message when multiple applications were found using getAppRegistrationByAppName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'&$select=id,displayName`) {
        return {
          value: [
            {
              id: appObjectId,
              displayName: appName
            },
            {
              id: secondAppObjectId,
              displayName: appName
            }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(entraApp.getAppRegistrationByAppName(appName, ['id', 'displayName']), Error(`Multiple apps with name '${appName}' found in Microsoft Entra ID. Found: ${appObjectId}, ${secondAppObjectId}.`));
  });
});