import assert from 'assert';
import sinon from 'sinon';
import auth, { AuthType, CertificateType, CloudType } from '../../../Auth.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { spo } from '../../../utils/spo.js';
import commands from '../commands.js';
import command from './identity-list.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import { CommandError } from '../../../Command.js';

describe(commands.LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const mockListResponse = [
    {
      "identityName": "alexw@contoso.com",
      "identityId": "028de82d-7fd9-476e-a9fd-be9714280ff3",
      "authType": "DeviceCode",
      "appId": "31359c7f-bd7e-475c-86db-fdb8c937548e",
      "appTenant": "common",
      "cloudType": "Public"
    },
    {
      "identityName": "Contoso Application",
      "identityId": "acd6df42-10a9-4315-8928-53334f1c9d01",
      "authType": "Secret",
      "appId": "39446e2e-5081-4887-980c-f285919fccca",
      "appTenant": "db308122-52f3-4241-af92-1734aa6e2e50",
      "cloudType": "Public"
    }
  ];

  before(() => {
    sinon.stub(auth, 'clearConnectionInfo').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    sinon.stub(auth as any, 'getConnectionInfoFromStorage').resolves({
      authType: AuthType.DeviceCode,
      active: true,
      identityName: 'alexw@contoso.com',
      identityId: '028de82d-7fd9-476e-a9fd-be9714280ff3',
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      certificateType: CertificateType.Unknown,
      accessTokens: {
        'https://graph.microsoft.com': {
          expiresOn: (new Date()).toISOString(),
          accessToken: 'abc'
        }
      },
      availableIdentities: [
        {
          authType: AuthType.DeviceCode,
          active: true,
          identityName: 'alexw@contoso.com',
          identityId: '028de82d-7fd9-476e-a9fd-be9714280ff3',
          appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
          tenant: 'common',
          cloudType: CloudType.Public,
          certificateType: CertificateType.Unknown,
          accessTokens: {
            'https://graph.microsoft.com': {
              expiresOn: (new Date()).toISOString(),
              accessToken: 'abc'
            }
          }
        },
        {
          authType: AuthType.Secret,
          active: true,
          identityName: 'Contoso Application',
          identityId: 'acd6df42-10a9-4315-8928-53334f1c9d01',
          appId: '39446e2e-5081-4887-980c-f285919fccca',
          tenant: 'db308122-52f3-4241-af92-1734aa6e2e50',
          cloudType: CloudType.Public,
          certificateType: CertificateType.Unknown,
          accessTokens: {
            'https://graph.microsoft.com': {
              expiresOn: (new Date()).toISOString(),
              accessToken: 'abc'
            }
          }
        }
      ]
    });
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
  });

  after(() => {
    sinon.restore();
    auth.service.logout();
  });


  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['identityName', 'authType']);
  });

  it('shows a list of signed in identities', async () => {
    await assert.doesNotReject(command.action(logger, { options: {} }));
    assert(loggerLogSpy.calledOnceWithExactly(mockListResponse));
  });

  it('fails with error when restoring auth information leads to error', async () => {
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));

    try {
      await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore(auth.restoreAuth);
    }
  });
});