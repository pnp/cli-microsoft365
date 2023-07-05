import assert from 'assert';
import chalk from 'chalk';
import fs from 'fs';
import path from 'path';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './project-permissions-grant.js';
import { spo } from '../../../../utils/spo.js';

describe(commands.PROJECT_PERMISSIONS_GRANT, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerStderrLogSpy: sinon.SinonSpy;
  const projectPath: string = 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react';
  const packagejsonContent = `{
    "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
    "solution": {
      "name": "hello-world-client-side-solution",
      "id": "ae661453-5f05-44f3-8312-66eb8f36f6fc",
      "version": "1.0.0.0",
      "includeClientSideAssets": true,
      "skipFeatureDeployment": true,
      "isDomainIsolated": false,
      "developer": {
        "name": "",
        "websiteUrl": "",
        "privacyUrl": "",
        "termsOfUseUrl": "",
        "mpnId": "Undefined-1.16.1"
      },
      "metadata": {
        "shortDescription": {
          "default": "hello world description"
        },
        "longDescription": {
          "default": "hello world description"
        },
        "screenshotPaths": [],
        "videoUrl": "",
        "categories": []
      },
      "features": [
        {
          "title": "Hello world Feature",
          "description": "The feature that activates elements of the hello world solution.",
          "id": "345cee0f-e4fb-464f-b649-20fc96b5f6aa",
          "version": "1.0.0.0"
        }
      ],
      "webApiPermissionRequests": [
        {
          "resource": "Microsoft Graph",
          "scope": "User.ReadBasic.All"
        }
      ]
    },
    "paths": {
      "zippedPackage": "solution/hello-worldsppkg"
    }
  }`;
  const grantResponse = {
    "ClientId": "90a2c08e-e786-4100-9ea9-36c261be6c0d",
    "ConsentType": "AllPrincipals",
    "IsDomainIsolated": false,
    "ObjectId": "jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM",
    "PackageName": null,
    "Resource": "Microsoft Graph",
    "ResourceId": "d6afc295-0a08-4777-a095-10bcb7e23903",
    "Scope": "User.ReadBasic.All"
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerStderrLogSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      (command as any).getProjectRoot,
      fs.existsSync,
      fs.readFileSync,
      spo.servicePrincipalGrant
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PROJECT_PERMISSIONS_GRANT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(null);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Couldn't find project root folder`, 1));
  });

  it('handles correctly when the package-solution.json file is not found', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), projectPath));

    sinon.stub(fs, 'existsSync').returns(false);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`The package-solution.json file could not be found`));
  });

  it('grant the specified permissions from the package-solution.json file', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(packagejsonContent);

    sinon.stub(spo, 'servicePrincipalGrant').resolves(grantResponse);

    await command.action(logger, {
      options: {
        debug: true
      }
    });
    assert(loggerLogSpy.calledWith(grantResponse));
  });

  it('shows warning when permission already exist', async () => {
    const grantExistError = {
      error: {
        message: 'An OAuth permission with the resource Microsoft Graph and scope User.ReadBasic.All already exists.Parameter name: permissionRequest'
      }
    };

    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(packagejsonContent);

    sinon.stub(spo, 'servicePrincipalGrant').rejects(grantExistError);

    await command.action(logger, {
      options: {
      }
    });
    assert.strictEqual(loggerStderrLogSpy.calledWith(chalk.yellow("An OAuth permission with the resource Microsoft Graph and scope User.ReadBasic.All already exists.Parameter name: permissionRequest")), true);
  });

  it('correctly handles error when something went wrong when granting permission', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(packagejsonContent);

    sinon.stub(spo, 'servicePrincipalGrant').rejects(new Error('Something went wrong'));

    await assert.rejects(command.action(logger, { options: {} } as any),
      'Something went wrong');
  });
});
