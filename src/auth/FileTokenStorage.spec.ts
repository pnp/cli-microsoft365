import * as assert from 'assert';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as sinon from 'sinon';
import { AuthType, CertificateType, Service } from '../Auth';
import Utils from '../Utils';
import { FileTokenStorage } from './FileTokenStorage';

describe('FileTokenStorage', () => {
  const fileStorage = new FileTokenStorage(FileTokenStorage.connectionInfoFilePath());

  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFile
    ]);
  });

  it(`stores MSAL cache in the user's home directory`, () => {
    assert.strictEqual(FileTokenStorage.msalCacheFilePath(), path.join(os.homedir(), '.cli-m365-msal.json'));
  });

  it('fails retrieving connection info from file if the token file doesn\'t exist', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    fileStorage
      .get()
      .then(() => {
        done('Expected fail but passed instead');
      }, (err) => {
        try {
          assert.strictEqual(err, 'File not found');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns connection info from file', (done) => {
    const tokensFile: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify(tokensFile));
    fileStorage
      .get()
      .then((connectionInfo) => {
        try {
          assert.strictEqual(connectionInfo, JSON.stringify(tokensFile));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('saves the connection info in the file when the file doesn\'t exist', (done) => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token as string; }).callsArgWith(3, null);
    fileStorage
      .set(JSON.stringify(expected))
      .then(() => {
        try {
          assert.strictEqual(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('saves the connection info in the file when the file is empty', (done) => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token as string; }).callsArgWith(3, null);
    fileStorage
      .set(JSON.stringify(expected))
      .then(() => {
        try {
          assert.strictEqual(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('saves the connection info in the file when the file contains an empty JSON object', (done) => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{}');
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token as string; }).callsArgWith(3, null);
    fileStorage
      .set(JSON.stringify(expected))
      .then(() => {
        try {
          assert.strictEqual(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('saves the connection info in the file when the file contains no access tokens', (done) => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{"accessTokens":{},"authType":0,"connected":false}');
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token as string; }).callsArgWith(3, null);
    fileStorage
      .set(JSON.stringify(expected))
      .then(() => {
        try {
          assert.strictEqual(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('adds the connection info to the file when the file contains access tokens', (done) => {
    const expected: Service = {
      accessTokens: {},
      appId: '31359c7f-bd7e-475c-86db-fdb8c937548e',
      tenant: 'common',
      authType: AuthType.DeviceCode,
      certificateType: CertificateType.Unknown,
      connected: false,
      logout: () => { }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
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
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token as string; }).callsArgWith(3, null);
    fileStorage
      .set(JSON.stringify(expected))
      .then(() => {
        try {
          assert.strictEqual(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('correctly handles error when writing to the file failed', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    sinon.stub(fs, 'writeFile').callsFake(() => { }).callsArgWith(3, { message: 'An error has occurred' });
    fileStorage
      .set('abc')
      .then(() => {
        done('Fail expected but passed instead');
      }, (err) => {
        try {
          assert.strictEqual(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('succeeds with removing if the token file doesn\'t exist', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    fileStorage
      .remove()
      .then(() => {
        done();
      }, () => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds with removing if the token file is empty', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    sinon.stub(fs, 'writeFile').callsFake(() => '').callsArgWith(3, null);
    fileStorage
      .remove()
      .then(() => {
        done();
      }, () => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds with removing if the token file contains empty JSON object', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{}');
    sinon.stub(fs, 'writeFile').callsFake(() => '').callsArgWith(3, null);
    fileStorage
      .remove()
      .then(() => {
        done();
      }, () => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds with removing if the token file contains no services', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({ services: {} }));
    sinon.stub(fs, 'writeFile').callsFake(() => { }).callsArgWith(3, null);
    fileStorage
      .remove()
      .then(() => {
        done();
      }, () => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds when connection info successfully removed from the token file', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      services: {
        'abc': 'def'
      }
    }));
    sinon.stub(fs, 'writeFile').callsFake(() => { }).callsArgWith(3, null);
    fileStorage
      .remove()
      .then(() => {
        try {
          done();
        }
        catch (e) {
          done(e);
        }
      }, () => {
        done('Pass expected but failed instead');
      });
  });

  it('correctly handles error when writing updated tokens to the token file', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      services: {
        'abc': 'def'
      }
    }));
    sinon.stub(fs, 'writeFile').callsFake(() => { }).callsArgWith(3, { message: 'An error has occurred' });
    fileStorage
      .remove()
      .then(() => {
        done('Fail expected but passed instead');
      }, (err) => {
        try {
          assert.strictEqual(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });
});