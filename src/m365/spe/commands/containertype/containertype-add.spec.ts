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
import { spo } from '../../../../utils/spo.js';
import { CommandError } from '../../../../Command.js';

describe(commands.CONTAINERTYPE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  let commandInfo: CommandInfo;

  const applicationId = 'f08575e2-36c4-407f-a891-eabae23f66bc';
  const containerName = 'New Container';
  const azureSubscriptionId = '440132a1-d85c-4e1a-b087-8c32ae189a4e';
  const region = 'West Europe';
  const resourceGroup = 'Resource Group';
  const adminUrl = 'https://contoso-admin.sharepoint.com';

  const errorMessageTooManyContainers = 'Maximum number of allowed Trial Container Types has been exceeded.';
  const errorMessageContainerAlreadyExists = `Container Type with Id c33cfee5-c9b6-0a2a-02ee-060693a57f37 already exists for Owning Application ${applicationId}`;

  const csomContainerAlreadyExistsError = `[{"SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.24621.12009","ErrorInfo":{"ErrorMessage":"${errorMessageContainerAlreadyExists}","ErrorValue":null,"TraceCorrelationId":"380514a1-50e2-8000-6705-77935619cb19","ErrorCode":-2146232832,"ErrorTypeName":"Microsoft.SharePoint.SPException"},"TraceCorrelationId":"380514a1-50e2-8000-6705-77935619cb19"}]`;
  const csomTrialError = `[{"SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.24621.12009","ErrorInfo":{"ErrorMessage":"${errorMessageTooManyContainers}","ErrorValue":null,"TraceCorrelationId":"df0214a1-d093-8000-4a04-15fbbddb489b","ErrorCode":-2146232832,"ErrorTypeName":"Microsoft.SharePoint.SPException"},"TraceCorrelationId":"df0214a1-d093-8000-4a04-15fbbddb489b"}]`;
  const csomOutput = `[{"SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.24621.12009","ErrorInfo":null,"TraceCorrelationId":"1dfe13a1-9073-8000-4a04-123c4e4bf45e"},4,{"_ObjectType_":"Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties","AzureSubscriptionId":"\/Guid(${azureSubscriptionId})\/","ContainerTypeId":"\/Guid(fafa338d-8bd5-04fe-3ea1-763649bff2df)\/","CreationDate":"3\u002f11\u002f2024 1:30:21 PM","DisplayName":"${containerName}","ExpiryDate":null,"IsBillingProfileRequired":true,"OwningAppId":"\/Guid(${applicationId})\/","OwningTenantId":"\/Guid(e1dd4023-a656-480a-8a0e-c1b1eec51e1d)\/","Region":"${region}","ResourceGroup":"${resourceGroup}","SPContainerTypeBillingClassification":0}]`;

  const jsonResponse = {
    AzureSubscriptionId: azureSubscriptionId,
    ContainerTypeId: 'fafa338d-8bd5-04fe-3ea1-763649bff2df',
    CreationDate: "3/11/2024 1:30:21 PM",
    DisplayName: containerName,
    ExpiryDate: null,
    IsBillingProfileRequired: true,
    OwningAppId: applicationId,
    OwningTenantId: "e1dd4023-a656-480a-8a0e-c1b1eec51e1d",
    Region: region,
    ResourceGroup: resourceGroup,
    SPContainerTypeBillingClassification: "Standard"
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getSpoAdminUrl').resolves(adminUrl);
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: adminUrl });

    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINERTYPE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a new trial Container Type', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return csomOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: containerName, applicationId: applicationId, trial: true, verbose: true } });
    const jsonResponseClone = { ...jsonResponse };
    jsonResponseClone.SPContainerTypeBillingClassification = "Trial";
    assert(loggerLogSpy.calledWith(jsonResponseClone));
  });

  it('creates a new standard Container Type', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return csomOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: containerName, applicationId: applicationId, region: region, azureSubscriptionId: azureSubscriptionId, resourceGroup: resourceGroup, verbose: true } });
    assert(loggerLogSpy.calledWith(jsonResponse));
  });

  it('throws an error when too many trial containers have been created', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return csomTrialError;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: containerName, applicationId: applicationId, trial: true, verbose: true } })
      , new CommandError(errorMessageTooManyContainers));
  });

  it('throws an error when container with application id has already been created', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return csomContainerAlreadyExistsError;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: containerName, applicationId: applicationId, azureSubscriptionId: azureSubscriptionId, resourceGroup: resourceGroup, region: region, verbose: true } })
      , new CommandError(errorMessageContainerAlreadyExists));
  });

  it('fails validation if the specified applicationId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { name: containerName, applicationId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the specified applicationId is a valid GUID and trial is specified', async () => {
    const actual = await command.validate({ options: { name: containerName, applicationId: applicationId, trial: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if azureSubscriptionId is not passed and trial is not passed', async () => {
    const actual = await command.validate({ options: { name: containerName, applicationId: applicationId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if resourceGroup is not passed and trial is not passed', async () => {
    const actual = await command.validate({ options: { name: containerName, applicationId: applicationId, azureSubscriptionId: azureSubscriptionId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if region is not passed and trial is not passed', async () => {
    const actual = await command.validate({ options: { name: containerName, applicationId: applicationId, azureSubscriptionId: azureSubscriptionId, resourceGroup: resourceGroup } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if azureSubscriptionId is not a valid GUID and trial is not passed', async () => {
    const actual = await command.validate({ options: { name: containerName, applicationId: applicationId, azureSubscriptionId: 'invalid', region: region, resourceGroup: resourceGroup } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all options are passed and it is not a trial', async () => {
    const actual = await command.validate({ options: { name: containerName, applicationId: applicationId, azureSubscriptionId: azureSubscriptionId, region: region, resourceGroup: resourceGroup } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
