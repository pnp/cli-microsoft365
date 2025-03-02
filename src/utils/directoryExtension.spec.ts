import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { directoryExtension } from './directoryExtension.js';
import { formatting } from './formatting.js';

describe('utils/directoryExtension', () => {
  const appObjectId = 'd6a8bfec-893d-46e4-88fd-7db5fcc0fa62';
  const name = 'extension_105be60b603845fea385e58772d9d630_GitHubWorkAccount';
  const invalidName = 'GitHubWorkAccount';
  const response = {
    "id": "522817ae-5c95-4243-96c1-f85231fcbc1f",
    "deletedDateTime": null,
    "appDisplayName": "ContosoApp",
    "dataType": "String",
    "isMultiValued": false,
    "isSyncedFromOnPremises": false,
    "name": "extension_105be60b603845fea385e58772d9d630_GitHubWorkAccount",
    "targetObjects": [
      "User"
    ]
  };
  const limitedResponse = {
    "id": "522817ae-5c95-4243-96c1-f85231fcbc1f",
    "name": "extension_105be60b603845fea385e58772d9d630_GitHubWorkAccount"
  };

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single directory extension by name using getDirectoryExtensionByName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties?$filter=name eq '${formatting.encodeQueryParameter(name)}'`) {
        return {
          value: [
            response
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await directoryExtension.getDirectoryExtensionByName(name, appObjectId);
    assert.deepStrictEqual(actual, {
      "id": "522817ae-5c95-4243-96c1-f85231fcbc1f",
      "deletedDateTime": null,
      "appDisplayName": "ContosoApp",
      "dataType": "String",
      "isMultiValued": false,
      "isSyncedFromOnPremises": false,
      "name": "extension_105be60b603845fea385e58772d9d630_GitHubWorkAccount",
      "targetObjects": [
        "User"
      ]
    });
  });

  it('correctly get single directory extension by name using getDirectoryExtensionByName with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties?$filter=name eq '${formatting.encodeQueryParameter(name)}'&$select=id,name`) {
        return {
          value: [
            limitedResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await directoryExtension.getDirectoryExtensionByName(name, appObjectId, ['id', 'name']);
    assert.deepStrictEqual(actual, {
      "id": "522817ae-5c95-4243-96c1-f85231fcbc1f",
      "name": "extension_105be60b603845fea385e58772d9d630_GitHubWorkAccount"
    });
  });

  it('throws error message when no directory extension was found using getDirectoryExtensionByName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/applications/${appObjectId}/extensionProperties?$filter=name eq '${formatting.encodeQueryParameter(invalidName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(directoryExtension.getDirectoryExtensionByName(invalidName, appObjectId)), Error(`The specified directory extension '${invalidName}' does not exist.`);
  });
});