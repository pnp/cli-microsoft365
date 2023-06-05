import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-export');

describe(commands.APP_EXPORT, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  const actualFilename = 'Power App.zip';
  const packageDisplayName = 'Power App';
  const packageDescription = 'Power App Description';
  const packageCreatedBy = 'John Doe';
  const packageSourceEnvironment = "Contoso";
  const path = 'c:/users/John/Documents';
  const environment = 'Default-cf409f12-a06f-426e-9955-20f5d7a31dd3';
  const appId = '11403f1a-de85-4b7d-97c9-020429876cb8';
  const listPackageResourcesResponse = {
    status: 'Succeeded',
    baseResourceIds: [
      '/providers/Microsoft.PowerApps/apps/11403f1a-de85-4b7d-97c9-020429876cb8'
    ],
    resources: {
      L1BST1ZJREVSUy9NSUNST1NPRlQuUE9XRVJBUFBTL0FQUFMvMTE0MDNGMUEtREU4NS00QjdELTk3QzktMDIwNDI5ODc2Q0I4: {
        id: `/providers/Microsoft.PowerApps/apps/${appId}`,
        name: appId,
        type: 'Microsoft.PowerApps/apps',
        creationType: 'New, Update',
        details: {
          displayName: 'App'
        },
        configurableBy: 'User',
        hierarchy: 'Root',
        dependsOn: []
      }
    }
  };

  const exportPackageResponse = {
    headers: {
      location: `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${environment}/packagingOperations/10fc880b-b11d-4fac-b842-386c66b869eb?api-version=2016-11-01`
    },
    data: {
      status: 'Running',
      details: {
        displayName: 'test',
        packageTelemetryId: '84b96b80-4593-4fb4-a35a-3cbe32be98ae'
      },
      resources: {
        '399ede40-1b69-4e28-ac8b-ab6899e617c7': {
          id: '/providers/Microsoft.PowerApps/apps/11403f1a-de85-4b7d-97c9-020429876cb8',
          name: appId,
          type: 'Microsoft.PowerApps/apps',
          status: 'Running',
          suggestedCreationType: 'Update',
          creationType: 'New, Update',
          details: {
            displayName: 'App'
          },
          configurableBy: 'User',
          hierarchy: 'Root',
          dependsOn: []
        }
      }
    }
  };

  const locationRunningResponse = {
    id: `/providers/Microsoft.BusinessAppPlatform/environments/${environment}/packagingOperations/10fc880b-b11d-4fac-b842-386c66b869eb`,
    type: 'Microsoft.BusinessAppPlatform/environments/packagingOperations',
    environmentName: environment,
    name: '10fc880b-b11d-4fac-b842-386c66b869eb',
    properties: {
      status: 'Running',
      details: {
        displayName: 'test',
        packageTelemetryId: '84b96b80-4593-4fb4-a35a-3cbe32be98ae'
      },
      resources: {
        '399ede40-1b69-4e28-ac8b-ab6899e617c7': {
          id: `/providers/Microsoft.PowerApps/apps/${appId}`,
          name: appId,
          type: 'Microsoft.PowerApps/apps',
          status: 'Running',
          suggestedCreationType: 'Update',
          creationType: 'New, Update',
          details: {
            displayName: 'App'
          },
          configurableBy: 'User',
          hierarchy: 'Root',
          dependsOn: []
        }
      }
    }
  };

  const locationSuccessResponse = {
    id: '/providers/Microsoft.BusinessAppPlatform/environments/Default-0cac6cda-2e04-4a3d-9c16-9c91470d7022/packagingOperations/10fc880b-b11d-4fac-b842-386c66b869eb',
    type: 'Microsoft.BusinessAppPlatform/environments/packagingOperations',
    environmentName: 'Default-0cac6cda-2e04-4a3d-9c16-9c91470d7022',
    name: '10fc880b-b11d-4fac-b842-386c66b869eb',
    properties: {
      status: 'Succeeded',
      packageLink: {
        value: 'https://bapfeblobprodam.blob.core.windows.net/20230303t000000z312d30573c1c498aad959706d35cb25e/Power_App_20230303140531.zip?sv=2018-03-28&sr=c&sig=X9TCSpygBeu7BmzLT7TrN0bni9Qg3VDF9xBp04eUOr0%3D&se=2023-03-03T15%3A05%3A31Z&sp=rl'
      },
      details: {
        displayName: packageDisplayName,
        createdTime: '2023-03-03T17:05:31.7937267Z',
        packageTelemetryId: '84b96b80-4593-4fb4-a35a-3cbe32be98ae'
      },
      resources: {
        '399ede40-1b69-4e28-ac8b-ab6899e617c7': {
          id: `/providers/Microsoft.PowerApps/apps/11403f1a-de85-4b7d-97c9-020429876cb8`,
          name: '11403f1a-de85-4b7d-97c9-020429876cb8',
          type: 'Microsoft.PowerApps/apps',
          status: 'Succeeded',
          suggestedCreationType: 'Update',
          creationType: 'New, Update',
          details: {
            displayName: 'App'
          },
          configurableBy: 'User',
          dependsOn: []
        }
      }
    }
  };

  const fileBlobResponse = {
    type: 'Buffer',
    data: [80, 75, 3, 4, 20, 0, 0, 0, 8, 0, 237, 115, 99, 86, 250, 76, 155, 216, 248, 3, 0, 0, 7, 8, 0, 0, 71, 0, 0, 0, 77, 105, 99, 114, 111, 115, 111, 102, 116, 46, 80, 111, 119, 101, 114, 65, 112, 112, 115, 47, 97, 112, 112, 115, 47, 49, 56, 48, 50, 54, 54, 51, 51, 48]
  };

  before(() => {
    (command as any).pollingInterval = 0;
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_EXPORT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('exports the specified App', async () => {
    let index = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === exportPackageResponse.headers.location) {
        if (index === 0) {
          index = 1;
          return locationRunningResponse;
        }
        else {
          return locationSuccessResponse;
        }
      }

      if (opts.url === locationSuccessResponse.properties.packageLink.value) {
        return fileBlobResponse;
      }

      throw 'invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${environment}/listPackageResources?api-version=2016-11-01`) {
        return listPackageResourcesResponse;
      }

      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${environment}/exportPackage?api-version=2016-11-01`) {
        return exportPackageResponse;
      }

      throw 'invalid request';
    });
    sinon.stub(fs, 'writeFileSync').returns();

    await assert.doesNotReject(command.action(logger, { options: { id: appId, environment: environment, packageDisplayName: packageDisplayName } }));
  });

  it('exports the specified App (debug)', async () => {
    let index = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === exportPackageResponse.headers.location) {
        if (index === 0) {
          index = 1;
          return locationRunningResponse;
        }
        else {
          return locationSuccessResponse;
        }
      }

      if (opts.url === locationSuccessResponse.properties.packageLink.value) {
        return fileBlobResponse;
      }

      throw 'invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${environment}/listPackageResources?api-version=2016-11-01`) {
        return listPackageResourcesResponse;
      }

      if (opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${environment}/exportPackage?api-version=2016-11-01`) {
        return exportPackageResponse;
      }

      throw 'invalid request';
    });
    sinon.stub(fs, 'writeFileSync').returns();

    await command.action(logger, { options: { verbose: true, id: appId, environment: environment, packageDisplayName: packageDisplayName, packageDescription: packageDescription, packageCreatedBy: packageCreatedBy, packageSourceEnvironment: packageSourceEnvironment, path: path } });
    assert(loggerLogToStderrSpy.calledWith(`File saved to path '${path}/${actualFilename}'`));
  });

  it('fails validation if the id is not a GUID', async () => {
    const actual = await command.validate({ options: { id: 'foo', environment: environment, packageDisplayName: packageDisplayName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified path doesn\'t exist', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = await command.validate({ options: { id: appId, environment: environment, packageDisplayName: packageDisplayName, path: '/path/not/found.zip' } }, commandInfo);
    sinonUtil.restore(fs.existsSync);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id, environment and packageDisplayName specified', async () => {
    const actual = await command.validate({ options: { id: appId, environment: environment, packageDisplayName: packageDisplayName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles API OData error', async () => {
    const error = {
      error: {
        message: `Something went wrong exporting the Microsoft Power App`
      }
    };

    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, { options: { id: appId, environment: environment, packageDisplayName: packageDisplayName } } as any),
      new CommandError(error.error.message));
  });
});
