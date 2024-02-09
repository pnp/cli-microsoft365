import assert from 'assert';
import fs from 'fs';
import os from 'os';
import path from 'path';
import sinon from 'sinon';
import { AuthType, CertificateType, CloudType, Service } from '../Auth.js';
import { sinonUtil } from '../utils/sinonUtil.js';
import { FileTokenStorage } from './FileTokenStorage.js';

describe('FileTokenStorage', () => {
  const fileStorage = new FileTokenStorage(FileTokenStorage.connectionInfoFilePath());

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFile
    ]);
  });

  it(`stores MSAL cache in the user's home directory`, () => {
    assert.strictEqual(FileTokenStorage.msalCacheFilePath(), path.join(os.homedir(), '.cli-m365-msal.json'));
  });

  it('fails retrieving connection info from file if the token file doesn\'t exist', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    try {
      await fileStorage.get();
      assert.fail('Expected fail but passed instead');
    }
    catch (err) {
      assert.strictEqual(err, 'File not found');
    }
  });

  it('returns connection info from file', async () => {
    const tokensFile: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify(tokensFile));
    const connectionInfo = await fileStorage.get();
    assert.strictEqual(connectionInfo, JSON.stringify(tokensFile));
  });

  it('saves the connection info in the file when the file doesn\'t exist', async () => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').returns(false);
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token as string; }).callsArgWith(3, null);
    await fileStorage.set(JSON.stringify(expected));
    assert.strictEqual(actual, JSON.stringify(expected));
  });

  it('saves the connection info in the file when the file is empty', async () => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('');
    sinon.stub(fs, 'writeFile').callsFake((_, token) => { actual = token as string; }).callsArgWith(3, null);
    await fileStorage.set(JSON.stringify(expected));
    assert.strictEqual(actual, JSON.stringify(expected));
  });

  it('saves the connection info in the file when the file contains an empty JSON object', async () => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('{}');
    sinon.stub(fs, 'writeFile').callsFake((_, token) => { actual = token as string; }).callsArgWith(3, null);

    await fileStorage.set(JSON.stringify(expected));
    assert.strictEqual(actual, JSON.stringify(expected));
  });

  it('saves the connection info in the file when the file contains no access tokens', async () => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('{"accessTokens":{},"authType":0,"connected":false}');
    sinon.stub(fs, 'writeFile').callsFake((_, token) => { actual = token as string; }).callsArgWith(3, null);

    await fileStorage.set(JSON.stringify(expected));
    assert.strictEqual(actual, JSON.stringify(expected));
  });

  it('adds the connection info to the file when the file contains access tokens', async () => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      cloudType: CloudType.Public,
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({
      accessTokens: {
        "https://contoso.sharepoint.com": {
          expiresOn: '123',
          value: '123'
        }
      },
      authType: AuthType.DeviceCode,
      connected: true,
      refreshToken: 'ref'
    }));
    sinon.stub(fs, 'writeFile').callsFake((_, token) => { actual = token as string; }).callsArgWith(3, null);
    await fileStorage.set(JSON.stringify(expected));
    assert.strictEqual(actual, JSON.stringify(expected));
  });

  it('correctly handles error when writing to the file failed', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    sinon.stub(fs, 'writeFile').returns().callsArgWith(3, { message: 'An error has occurred' });
    try {
      await fileStorage.set('abc');
      assert.fail('Expected fail but passed instead');
    }
    catch (err) {
      assert.strictEqual(err, 'An error has occurred');
    }
  });

  it('succeeds with removing if the token file doesn\'t exist', async () => {
    const writeStub = sinon.stub(fs, 'writeFile');
    sinon.stub(fs, 'existsSync').returns(false);

    await fileStorage.remove();
    assert(writeStub.notCalled);
  });

  it('succeeds with removing if the token file is empty', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('');
    const writeStub = sinon.stub(fs, 'writeFile').returns().callsArgWith(3, null);

    await fileStorage.remove();
    assert(writeStub.called);
  });

  it('succeeds with removing if the token file contains empty JSON object', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('{}');
    const writeStub = sinon.stub(fs, 'writeFile').returns().callsArgWith(3, null);

    await fileStorage.remove();
    assert(writeStub.called);
  });

  it('succeeds with removing if the token file contains no services', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({ services: {} }));
    const writeStub = sinon.stub(fs, 'writeFile').returns().callsArgWith(3, null);

    await fileStorage.remove();
    assert(writeStub.called);
  });

  it('succeeds when connection info successfully removed from the token file', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({
      services: {
        'abc': 'def'
      }
    }));
    const writeStub = sinon.stub(fs, 'writeFile').returns().callsArgWith(3, null);

    await fileStorage.remove();
    assert(writeStub.called);
  });

  it('correctly handles error when writing updated tokens to the token file', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({
      services: {
        'abc': 'def'
      }
    }));
    sinon.stub(fs, 'writeFile').returns().callsArgWith(3, { message: 'An error has occurred' });
    try {
      await fileStorage.remove();
      assert.fail('Expected fail but passed instead');
    }
    catch (err) {
      assert.strictEqual(err, 'An error has occurred');
    }
  });
});