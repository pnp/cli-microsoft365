import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './containertype-list.js';
import { CommandError } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { accessToken } from '../../../../utils/accessToken.js';

const containerTypeData = [
  {
    id: 'de988700-d700-020e-0a00-0831f3042f00',
    name: 'Container Type 1',
    owningAppId: '11335700-9a00-4c00-84dd-0c210f203f00',
    billingClassification: 'trial',
    createdDateTime: '01/20/2025',
    expirationDateTime: '02/20/2025',
    etag: 'RVRhZw==',
    settings: {
      urlTemplate: 'https://app.contoso.com/redirect?tenant={tenant-id}&drive={drive-id}&folder={folder-id}&item={item-id}',
      isDiscoverabilityEnabled: true,
      isSearchEnabled: true,
      isItemVersioningEnabled: true,
      itemMajorVersionLimit: 50,
      maxStoragePerContainerInBytes: 104857600,
      isSharingRestricted: false,
      consumingTenantOverridables: ''
    }
  },
  {
    id: '88aeae-d700-020e-0a00-0831f3042f01',
    name: 'Container Type 2',
    owningAppId: '33225700-9a00-4c00-84dd-0c210f203f01',
    billingClassification: 'standard',
    createdDateTime: '01/20/2025',
    expirationDateTime: null,
    etag: 'RVRhZw==',
    settings: {
      urlTemplate: '',
      isDiscoverabilityEnabled: true,
      isSearchEnabled: true,
      isItemVersioningEnabled: false,
      itemMajorVersionLimit: 100,
      maxStoragePerContainerInBytes: 104857600,
      isSharingRestricted: false,
      consumingTenantOverridables: ''
    }
  }
];

describe(commands.CONTAINERTYPE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').withArgs('delegated').returns();
    auth.connection.active = true;
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
      odata.getAllItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINERTYPE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'owningAppId']);
  });

  it('retrieves list of container type', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === 'https://graph.microsoft.com/beta/storage/fileStorage/containerTypes') {
        return containerTypeData;
      }

      throw 'Invalid GET request ' + url;
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledOnceWith(containerTypeData));
  });

  it('correctly handles error when retrieving container types', async () => {
    const error = 'An error has occurred';
    sinon.stub(odata, 'getAllItems').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true
      }
    }), new CommandError('An error has occurred'));
  });
});