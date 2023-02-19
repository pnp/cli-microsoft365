import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { telemetry } from '../../../../telemetry';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { Cli } from '../../../../cli/Cli';
import * as SpoServicePrincipalGrantAddCommand from '../../../spo/commands/serviceprincipal/serviceprincipal-grant-add';
import chalk = require('chalk');
const command: Command = require('./project-permissions-grant');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerStderrLogSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      (command as any).getProjectRoot,
      fs.existsSync,
      fs.readFileSync,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PROJECT_PERMISSIONS_GRANT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => null);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Couldn't find project root folder`, 1));
  });

  it('handles correctly when the package-solution.json file is not found', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));

    sinon.stub(fs, 'existsSync').callsFake(_ => false);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`The package-solution.json file could not be found`));
  });

  it('grant the specified permissions from the package-solution.json file', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => packagejsonContent);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoServicePrincipalGrantAddCommand) {
        return ({
          stdout: `{ "ClientId": "90a2c08e-e786-4100-9ea9-36c261be6c0d", "ConsentType": "AllPrincipals", "IsDomainIsolated": false, "ObjectId": "jsCikIbnAEGeqTbCYb5sDZXCr9YICndHoJUQvLfiOQM", "PackageName": null, "Resource": "Microsoft Graph", "ResourceId": "d6afc295-0a08-4777-a095-10bcb7e23903", "Scope": "User.ReadBasic.All"}`
        });
      }

      throw new CommandError('Unknown case');
    });

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
      },
      stderr: ''
    };

    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => packagejsonContent);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoServicePrincipalGrantAddCommand) {
        throw grantExistError;
      }

      throw new CommandError('Unknown case');
    });

    await command.action(logger, {
      options: {
      }
    });
    assert.strictEqual(loggerStderrLogSpy.calledWith(chalk.yellow("An OAuth permission with the resource Microsoft Graph and scope User.ReadBasic.All already exists.Parameter name: permissionRequest")), true);
  });

  it('correctly handles error when something went wrong when granting permission', async () => {
    sinon.stub(fs, 'existsSync').callsFake(_ => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => packagejsonContent);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoServicePrincipalGrantAddCommand) {
        throw 'Something went wrong';
      }

      throw new CommandError('Unknown case');
    });

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Something went wrong`));
  });


});
