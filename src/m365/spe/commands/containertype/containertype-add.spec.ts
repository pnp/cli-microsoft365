import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './containertype-add.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { CommandError } from '../../../../Command.js';
import { z } from 'zod';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.CONTAINERTYPE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandOptionsSchema: z.ZodTypeAny;

  let commandInfo: CommandInfo;

  const applicationId = 'f08575e2-36c4-407f-a891-eabae23f66bc';
  const containerName = 'New Container';

  const containerTypeResponse = {
    id: 'de988700-d700-020e-0a00-0831f3042f00',
    name: 'Test Trial Container',
    owningAppId: '11335700-9a00-4c00-84dd-0c210f203f00',
    billingClassification: 'trial',
    billingStatus: 'valid',
    createdDateTime: '01/20/2025',
    expirationDateTime: '02/20/2025',
    etag: 'RVRhZw==',
    settings: {
      urlTemplate: '',
      isDiscoverabilityEnabled: true,
      isSearchEnabled: true,
      isItemVersioningEnabled: true,
      itemMajorVersionLimit: 50,
      maxStoragePerContainerInBytes: 104857600,
      isSharingRestricted: false,
      consumingTenantOverridables: 'isSearchEnabled,itemMajorVersionLimit'
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').withArgs('delegated').resolves();

    auth.connection.active = true;
    auth.connection.appId = 'a0de833a-3629-489a-8fc8-4dd0c431878c';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.appId = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINERTYPE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the specified applicationId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, appId: 'invalid' });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation if the specified applicationId is a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, appId: applicationId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if itemMajorVersionLimit is a negative number', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, itemMajorVersionLimit: -1 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if itemMajorVersionLimit is float number', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, itemMajorVersionLimit: 1.5 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if itemMajorVersionLimit is specified and isItemVersioningEnabled is false', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, itemMajorVersionLimit: 10, isItemVersioningEnabled: false });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if maxStoragePerContainerInBytes is a negative number', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, maxStoragePerContainerInBytes: -1 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if maxStoragePerContainerInBytes is float number', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, maxStoragePerContainerInBytes: 1.5 });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when billingType has an invalid value', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, billingType: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when sharingCapability has an invalid value', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, sharingCapability: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if consumingTenantOverridables contains an invalid value', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, consumingTenantOverridables: 'isDiscoverabilityEnabled, invalid, isItemVersioningEnabled' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if all options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, appId: applicationId, billingType: 'trial', consumingTenantOverridables: 'isItemVersioningEnabled,isDiscoverabilityEnabled', isDiscoverabilityEnabled: true, isItemVersioningEnabled: true, isSearchEnabled: true, isSharingRestricted: true, itemMajorVersionLimit: 23, maxStoragePerContainerInBytes: 12345, sharingCapability: 'disabled', urlTemplate: 'https://microsoft.com' });
    assert.strictEqual(actual.success, true);
  });

  it('creates a container type and outputs a result', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes`) {
        return containerTypeResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(containerTypeResponse));
  });

  it('creates a container type correctly', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes`) {
        return containerTypeResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, appId: applicationId, billingType: 'trial', consumingTenantOverridables: 'isItemVersioningEnabled, isDiscoverabilityEnabled', isDiscoverabilityEnabled: true, isItemVersioningEnabled: true, isSearchEnabled: true, isSharingRestricted: true, itemMajorVersionLimit: 23, maxStoragePerContainerInBytes: 12345, sharingCapability: 'disabled', urlTemplate: 'https://microsoft.com' } });
    assert.deepStrictEqual(postStub.firstCall.args[0]?.data, {
      name: containerName,
      owningAppId: applicationId,
      billingClassification: 'trial',
      settings: {
        consumingTenantOverridables: 'isItemVersioningEnabled,isDiscoverabilityEnabled',
        isDiscoverabilityEnabled: true,
        isItemVersioningEnabled: true,
        isSearchEnabled: true,
        isSharingRestricted: true,
        itemMajorVersionLimit: 23,
        maxStoragePerContainerInBytes: 12345,
        sharingCapability: 'disabled',
        urlTemplate: 'https://microsoft.com'
      }
    });
  });

  it('creates a container type correctly for current app', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes`) {
        return containerTypeResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, billingType: 'directToCustomer', consumingTenantOverridables: 'itemMajorVersionLimit,urlTemplate, maxStoragePerContainerInBytes', isDiscoverabilityEnabled: true, isItemVersioningEnabled: true, isSearchEnabled: false, isSharingRestricted: true, itemMajorVersionLimit: 23, maxStoragePerContainerInBytes: 12345, sharingCapability: 'disabled', urlTemplate: 'https://microsoft.com' } });
    assert.deepStrictEqual(postStub.firstCall.args[0]?.data, {
      name: containerName,
      owningAppId: 'a0de833a-3629-489a-8fc8-4dd0c431878c',
      billingClassification: 'directToCustomer',
      settings: {
        consumingTenantOverridables: 'itemMajorVersionLimit,urlTemplate,maxStoragePerContainerInBytes',
        isDiscoverabilityEnabled: true,
        isItemVersioningEnabled: true,
        isSearchEnabled: false,
        isSharingRestricted: true,
        itemMajorVersionLimit: 23,
        maxStoragePerContainerInBytes: 12345,
        sharingCapability: 'disabled',
        urlTemplate: 'https://microsoft.com'
      }
    });
  });

  it('throws an error when container type with application id has already been created', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        code: 'invalidRequest',
        message: 'Invalid request',
        innerError: {
          'request-id': '56ad0703-c9a5-4413-a86f-5617ee07f903',
          'client-request-id': '56ad0703-c9a5-4413-a86f-5617ee07f903'
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { name: containerName } }),
      new CommandError('Invalid request'));
  });
});
