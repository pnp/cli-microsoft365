import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../Auth.js';
import { CommandError } from '../Command.js';
import { accessToken } from '../utils/accessToken.js';
import { sinonUtil } from './sinonUtil.js';

describe('utils/accessToken', () => {

  before(() => {
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('isAppOnlyAccessToken returns undefined when access token is undefined', () => {
    const actual = accessToken.isAppOnlyAccessToken(undefined as any);
    assert.strictEqual(actual, undefined);
  });

  it('isAppOnlyAccessToken returns undefined when access token is empty', () => {
    const actual = accessToken.isAppOnlyAccessToken('');
    assert.strictEqual(actual, undefined);
  });

  it('isAppOnlyAccessToken returns undefined when access token is invalid', () => {
    const actual = accessToken.isAppOnlyAccessToken('abc.def');
    assert.strictEqual(actual, undefined);
  });

  it('isAppOnlyAccessToken returns undefined when non base64 access token passed', () => {
    const actual = accessToken.isAppOnlyAccessToken('abc.def.ghi');
    assert.strictEqual(actual, undefined);
  });

  it('isAppOnlyAccessToken returns true when access token is valid', () => {
    const jwt = JSON.stringify({
      idtyp: 'app'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const token = `abc.${jwt64}.def`;
    const actual = accessToken.isAppOnlyAccessToken(token);
    assert.strictEqual(actual, true);
  });

  it('shows app display name as connected-as for app-only auth', () => {
    const jwt = JSON.stringify({
      app_displayname: 'CLI for Microsoft 365 Contoso'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const token = `abc.${jwt64}.def`;
    const actual = accessToken.getUserNameFromAccessToken(token);
    assert.strictEqual(actual, 'CLI for Microsoft 365 Contoso');
  });

  it('returns empty user name when access token is undefined', () => {
    const actual = accessToken.getUserNameFromAccessToken(undefined as any);
    assert.strictEqual(actual, '');
  });

  it('returns empty user name when empty access token passed', () => {
    const actual = accessToken.getUserNameFromAccessToken('');
    assert.strictEqual(actual, '');
  });

  it('returns empty user name when invalid access token passed', () => {
    const actual = accessToken.getUserNameFromAccessToken('abc.def.ghi');
    assert.strictEqual(actual, '');
  });

  it('shows tenant id from valid access token', () => {
    const jwt = JSON.stringify({
      tid: 'de349bc7-1aeb-4506-8cb3-98db021cadb4'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const token = `abc.${jwt64}.def`;
    const actual = accessToken.getTenantIdFromAccessToken(token);
    assert.strictEqual(actual, 'de349bc7-1aeb-4506-8cb3-98db021cadb4');
  });

  it('returns empty tenant id when access token is undefined', () => {
    const actual = accessToken.getTenantIdFromAccessToken(undefined as any);
    assert.strictEqual(actual, '');
  });

  it('returns empty tenant id when empty access token passed', () => {
    const actual = accessToken.getTenantIdFromAccessToken('');
    assert.strictEqual(actual, '');
  });

  it('returns empty tenant id when invalid access token passed', () => {
    const actual = accessToken.getTenantIdFromAccessToken('abc.def.ghi');
    assert.strictEqual(actual, '');
  });

  it('shows user id from valid access token', () => {
    const jwt = JSON.stringify({
      oid: 'de349bc7-1aeb-4506-8cb3-98db021cadb4'
    });
    const jwt64 = Buffer.from(jwt).toString('base64');
    const token = `abc.${jwt64}.def`;
    const actual = accessToken.getUserIdFromAccessToken(token);
    assert.strictEqual(actual, 'de349bc7-1aeb-4506-8cb3-98db021cadb4');
  });

  it('returns empty userd id when access token is undefined', () => {
    const actual = accessToken.getUserIdFromAccessToken(undefined as any);
    assert.strictEqual(actual, '');
  });

  it('returns empty user id when empty access token passed', () => {
    const actual = accessToken.getUserIdFromAccessToken('');
    assert.strictEqual(actual, '');
  });

  it('returns empty user id when invalid access token passed', () => {
    const actual = accessToken.getUserIdFromAccessToken('abc.def.ghi');
    assert.strictEqual(actual, '');
  });

  it('returns empty user id when incomplete access token passed', () => {
    const actual = accessToken.getUserIdFromAccessToken('abc.def');
    assert.strictEqual(actual, '');
  });

  it('decodes access token', async () => {
    const decodedAccessToken = accessToken.getDecodedAccessToken('eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNTE2MjM5MDIyfQ.SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQssw5c');
    assert.deepStrictEqual(decodedAccessToken, {
      header: {
        alg: "HS256",
        typ: "JWT"
      },
      payload: {
        sub: "1234567890",
        name: "John Doe",
        iat: 1516239022
      }
    });
  });

  it('asserts delegated access token correctly', () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    accessToken.assertAccessTokenType('delegated');
  });

  it('throws error when trying to assert delegated access token when no token available', () => {
    const currentAccessTokens = auth.connection.accessTokens;
    auth.connection.accessTokens = {};
    assert.throws(() => accessToken.assertAccessTokenType('delegated'), new CommandError('No access token found.'));
    auth.connection.accessTokens = currentAccessTokens;
  });

  it('throws error when trying to assert delegated access token with application only token', () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    assert.throws(() => accessToken.assertAccessTokenType('delegated'), new CommandError('This command requires delegated permissions.'));
  });

  it('returns empty array for getScopesFromAccessToken when token is empty', () => {
    assert.deepStrictEqual(accessToken.getScopesFromAccessToken(''), []);
  });

  it('returns empty array for getScopesFromAccessToken when token has invalid format', () => {
    assert.deepStrictEqual(accessToken.getScopesFromAccessToken('invalid-token'), []);
  });

  it('returns empty array for getScopesFromAccessToken when token has no scp claim', () => {
    const payload = Buffer.from(JSON.stringify({ aud: 'https://graph.microsoft.com' })).toString('base64');
    const token = `header.${payload}.signature`;
    assert.deepStrictEqual(accessToken.getScopesFromAccessToken(token), []);
  });

  it('returns scopes from access token with single scope', () => {
    const payload = Buffer.from(JSON.stringify({ scp: 'Calendars.Read' })).toString('base64');
    const token = `header.${payload}.signature`;
    assert.deepStrictEqual(accessToken.getScopesFromAccessToken(token), ['Calendars.Read']);
  });

  it('returns scopes from access token with multiple scopes', () => {
    const payload = Buffer.from(JSON.stringify({ scp: 'Calendars.Read Calendars.Read.Shared Mail.Read' })).toString('base64');
    const token = `header.${payload}.signature`;
    assert.deepStrictEqual(accessToken.getScopesFromAccessToken(token), ['Calendars.Read', 'Calendars.Read.Shared', 'Mail.Read']);
  });

  it('returns empty array for getScopesFromAccessToken when token payload is not valid JSON', () => {
    const payload = Buffer.from('not-json').toString('base64');
    const token = `header.${payload}.signature`;
    assert.deepStrictEqual(accessToken.getScopesFromAccessToken(token), []);
  });

  it('reads access token from JSON token file using access_token property', () => {
    const jwtPayload = Buffer.from(JSON.stringify({
      exp: 1893456000,
      oid: '028de82d-7fd9-476e-a9fd-be9714280ff3'
    })).toString('base64');
    const jwt = `abc.${jwtPayload}.def`;
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({ access_token: jwt }));

    const actual = accessToken.readAccessTokenFromFile('/tmp/token.json');

    assert.strictEqual(actual.accessToken, jwt);
    assert.strictEqual((actual.expiresOn as Date).getTime(), 1893456000000);
  });

  it('reads access token from JSON token file using accessToken property', () => {
    const jwtPayload = Buffer.from(JSON.stringify({
      exp: 1893456000
    })).toString('base64');
    const jwt = `abc.${jwtPayload}.def`;
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({ accessToken: jwt }));

    const actual = accessToken.readAccessTokenFromFile('/tmp/token.json');

    assert.strictEqual(actual.accessToken, jwt);
    assert.strictEqual((actual.expiresOn as Date).getTime(), 1893456000000);
  });

  it('reads access token from plain text token file', () => {
    const jwtPayload = Buffer.from(JSON.stringify({
      exp: 1893456000
    })).toString('base64');
    const jwt = `abc.${jwtPayload}.def`;
    sinon.stub(fs, 'readFileSync').returns(jwt);

    const actual = accessToken.readAccessTokenFromFile('/tmp/token.txt');

    assert.strictEqual(actual.accessToken, jwt);
    assert.strictEqual((actual.expiresOn as Date).getTime(), 1893456000000);
  });

  it('reads expires_on from JSON token file', () => {
    const jwtPayload = Buffer.from(JSON.stringify({ oid: '028de82d-7fd9-476e-a9fd-be9714280ff3' })).toString('base64');
    const jwt = `abc.${jwtPayload}.def`;
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({
      access_token: jwt,
      expires_on: '1893456000'
    }));

    const actual = accessToken.readAccessTokenFromFile('/tmp/token.json');

    assert.strictEqual(actual.accessToken, jwt);
    assert.strictEqual((actual.expiresOn as Date).getTime(), 1893456000000);
  });

  it('reads string expiresOn from JSON token file', () => {
    const jwtPayload = Buffer.from(JSON.stringify({ oid: '028de82d-7fd9-476e-a9fd-be9714280ff3' })).toString('base64');
    const jwt = `abc.${jwtPayload}.def`;
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({
      accessToken: jwt,
      expiresOn: '2030-01-01T00:00:00.000Z'
    }));

    const actual = accessToken.readAccessTokenFromFile('/tmp/token.json');

    assert.strictEqual((actual.expiresOn as Date).getTime(), new Date('2030-01-01T00:00:00.000Z').getTime());
  });

  it('reads numeric expiresOn from JSON token file', () => {
    const jwtPayload = Buffer.from(JSON.stringify({ oid: '028de82d-7fd9-476e-a9fd-be9714280ff3' })).toString('base64');
    const jwt = `abc.${jwtPayload}.def`;
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify({
      accessToken: jwt,
      expiresOn: 1893456000
    }));

    const actual = accessToken.readAccessTokenFromFile('/tmp/token.json');

    assert.strictEqual((actual.expiresOn as Date).getTime(), 1893456000000);
  });

  it('throws error when token file does not contain a valid access token', () => {
    sinon.stub(fs, 'readFileSync').returns('not-a-token');

    assert.throws(() => accessToken.readAccessTokenFromFile('/tmp/token.txt'), (err: any) => {
      return err instanceof CommandError && err.message === 'Token file does not contain a valid access token.';
    });
  });

  it('throws error when JSON token file does not contain an access token property', () => {
    sinon.stub(fs, 'readFileSync').returns('{}');

    assert.throws(() => accessToken.readAccessTokenFromFile('/tmp/token.json'), (err: any) => {
      return err instanceof CommandError && err.message === 'Token file does not contain a valid access token.';
    });
  });

  it('reads access token from JSON token file containing a string value', () => {
    const jwtPayload = Buffer.from(JSON.stringify({
      exp: 1893456000
    })).toString('base64');
    const jwt = `abc.${jwtPayload}.def`;
    sinon.stub(fs, 'readFileSync').returns(JSON.stringify(jwt));

    const actual = accessToken.readAccessTokenFromFile('/tmp/token.json');

    assert.strictEqual(actual.accessToken, jwt);
    assert.strictEqual((actual.expiresOn as Date).getTime(), 1893456000000);
  });

  it('returns null expiresOn when token payload cannot be parsed', () => {
    const jwt = 'abc.!!!.def';
    sinon.stub(fs, 'readFileSync').returns(jwt);

    const actual = accessToken.readAccessTokenFromFile('/tmp/token.txt');

    assert.strictEqual(actual.accessToken, jwt);
    assert.strictEqual(actual.expiresOn, null);
  });
});