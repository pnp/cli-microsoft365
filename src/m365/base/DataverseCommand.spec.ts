import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../Auth';
import request from '../../request';
import { sinonUtil } from '../../utils';


import DataverseCommand from './DataverseCommand';

class MockCommand extends DataverseCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(): Promise<void> {
    return Promise.resolve();
  }

  public commandHelp(): void {
  }
}

describe('DataverseCommand', () => {
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    auth.service.connected = true;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('defines correct resource', () => {
    const cmd = new MockCommand();
    assert.strictEqual((cmd as any).resource, 'https://api.bap.microsoft.com');
  });

  it('returns correct dynamics url as admin', (done) => {
    const cmd = new MockCommand();
    const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(envResponse);
        }
      }

      return Promise.reject('Invalid request');
    });

    (cmd as any).getDynamicsInstance('someRandomGuid', true).then((instanceUrl: string) => {
      try {
        assert(instanceUrl === 'https://contoso-dev.api.crm4.dynamics.com');
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('returns correct dynamics url', (done) => {
    const cmd = new MockCommand();
    const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(envResponse);
        }
      }

      return Promise.reject('Invalid request');
    });

    (cmd as any).getDynamicsInstance('someRandomGuid', false).then((instanceUrl: string) => {
      try {
        assert(instanceUrl === 'https://contoso-dev.api.crm4.dynamics.com');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

});
