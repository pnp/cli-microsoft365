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
import command from './containertype-list.js';
import { spe } from '../../../../utils/spe.js';
import { CommandError } from '../../../../Command.js';

const containerTypeData = [{
  AzureSubscriptionId: 'f08575e2-36c4-407f-a891-eabae23f66bc',
  ContainerTypeId: 'c33cfee5-c9b6-0a2a-02ee-060693a57f37',
  CreationDate: '3/11/2024 2:38:56 PM',
  DisplayName: 'standard container',
  ExpiryDate: '3/11/2028 2:38:56 PM',
  IsBillingProfileRequired: true,
  OwningAppId: '1b3b8660-9a44-4a7c-9c02-657f3ff5d5ac',
  OwningTenantId: 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d',
  Region: 'West Europe',
  ResourceGroup: 'Standard group',
  SPContainerTypeBillingClassification: 'Standard'
},
{
  AzureSubscriptionId: 'f08575e2-36c4-407f-a891-eabae23f66bc',
  ContainerTypeId: 'a33cfee5-c9b6-0a2a-02ee-060693a57f37',
  CreationDate: '3/11/2024 2:38:56 PM',
  DisplayName: 'trial container',
  ExpiryDate: '3/11/2028 2:38:56 PM',
  IsBillingProfileRequired: true,
  OwningAppId: '1b3b8660-9a44-4a7c-9c02-657f3ff5d5ac',
  OwningTenantId: 'e1dd4023-a656-480a-8a0e-c1b1eec51e1d',
  Region: 'West Europe',
  ResourceGroup: 'Standard group',
  SPContainerTypeBillingClassification: 'Standard'
}];

describe(commands.CONTAINERTYPE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
      request.post,
      spe.getAllContainerTypes
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINERTYPE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['ContainerTypeId', 'DisplayName', 'OwningAppId']);
  });

  it('retrieves list of container type', async () => {
    sinon.stub(spe, 'getAllContainerTypes').resolves(containerTypeData);

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledOnceWith([
      {
        _ObjectType_: 'Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties',
        AzureSubscriptionId: '/Guid(f08575e2-36c4-407f-a891-eabae23f66bc)/',
        ContainerTypeId: '/Guid(c33cfee5-c9b6-0a2a-02ee-060693a57f37)/',
        CreationDate: '3/11/2024 2:38:56 PM',
        DisplayName: 'standard container',
        ExpiryDate: '3/11/2028 2:38:56 PM',
        IsBillingProfileRequired: true,
        OwningAppId: '/Guid(1b3b8660-9a44-4a7c-9c02-657f3ff5d5ac)/',
        OwningTenantId: '/Guid(e1dd4023-a656-480a-8a0e-c1b1eec51e1d)/',
        Region: 'West Europe',
        ResourceGroup: 'Standard group',
        SPContainerTypeBillingClassification: 'Standard'
      },
      {
        _ObjectType_: 'Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties',
        AzureSubscriptionId: '/Guid(f08575e2-36c4-407f-a891-eabae23f66bc)/',
        ContainerTypeId: '/Guid(a33cfee5-c9b6-0a2a-02ee-060693a57f37)/',
        CreationDate: '3/11/2024 2:38:56 PM',
        DisplayName: 'trial container',
        ExpiryDate: '3/11/2028 2:38:56 PM',
        IsBillingProfileRequired: true,
        OwningAppId: '/Guid(1b3b8660-9a44-4a7c-9c02-657f3ff5d5ac)/',
        OwningTenantId: '/Guid(e1dd4023-a656-480a-8a0e-c1b1eec51e1d)/',
        Region: 'West Europe',
        ResourceGroup: 'Standard group',
        SPContainerTypeBillingClassification: 'Standard'
      }
    ]));
  });

  it('correctly handles error when retrieving container types', async () => {
    const error = 'An error has occurred';
    sinon.stub(spe, 'getAllContainerTypes').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true
      }
    }), new CommandError('An error has occurred'));
  });
});