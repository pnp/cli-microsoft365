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
import { spo } from '../../../../utils/spo.js';
import { CommandError } from '../../../../Command.js';

const containerTypedata = [{
  "AzureSubscriptionId": "/Guid(f08575e2-36c4-407f-a891-eabae23f66bc)",
  "ContainerTypeId": "/Guid(c33cfee5-c9b6-0a2a-02ee-060693a57f37)",
  "CreationDate": "3/11/2024 2:38:56 PM",
  "DisplayName": "standard container",
  "ExpiryDate": "3/11/2028 2:38:56 PM",
  "IsBillingProfileRequired": true,
  "OwningAppId": "/Guid(1b3b8660-9a44-4a7c-9c02-657f3ff5d5ac)",
  "OwningTenantId": "/Guid(e1dd4023-a656-480a-8a0e-c1b1eec51e1d)",
  "Region": "West Europe",
  "ResourceGroup": "Standard group",
  "SPContainerTypeBillingClassification": "Standard"
},
{
  "AzureSubscriptionId": "/Guid(f08575e2-36c4-407f-a891-eabae23f66bc)",
  "ContainerTypeId": "/Guid(c33cfee5-c9b6-0a2a-02ee-060693a57f37)",
  "CreationDate": "3/11/2024 2:38:56 PM",
  "DisplayName": "trial container",
  "ExpiryDate": "3/11/2028 2:38:56 PM",
  "IsBillingProfileRequired": true,
  "OwningAppId": "/Guid(1b3b8660-9a44-4a7c-9c02-657f3ff5d5ac)",
  "OwningTenantId": "/Guid(e1dd4023-a656-480a-8a0e-c1b1eec51e1d)",
  "Region": "West Europe",
  "ResourceGroup": "Standard group",
  "SPContainerTypeBillingClassification": "Standard"
}];

describe(commands.CONTAINERTYPE_LIST, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      spo.getAllContainerTypes
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
    sinon.stub(spo, 'getAllContainerTypes').resolves(containerTypedata);
    await command.action(logger, { options: { debug: true } });
  });

  it('correctly handles error when retrieving container types', async () => {
    const error = 'An error has occurred';
    sinon.stub(spo, 'getAllContainerTypes').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true
      }
    }), new CommandError('An error has occurred'));
  });
});