import * as assert from 'assert';

describe('Config', () => {
  before(() => {
    delete require.cache[require.resolve('./config')];
  });

  afterEach(() => {
    delete require.cache[require.resolve('./config')];
  });

  it('returns process.env CLIMICROSOFT365_TENANT value', () => {
    process.env.CLIMICROSOFT365_TENANT = 'tenant123';

    const config = require('./config');
    assert.strictEqual(config.default.tenant, 'tenant123');
  });

  it('returns process.env CLIMICROSOFT365_AADAPPID value', () => {
    process.env.CLIMICROSOFT365_AADAPPID = 'appId123';

    const config = require('./config');
    assert.strictEqual(config.default.cliAadAppId, 'appId123');
  });

  it('returns default value since env CLIMICROSOFT365_TENANT not present', () => {
    delete process.env.CLIMICROSOFT365_TENANT;

    const config = require('./config');
    assert.strictEqual(config.default.tenant, 'common');
  });

  it('returns default value since env CLIMICROSOFT365_AADAPPID not present', () => {
    delete process.env.CLIMICROSOFT365_AADAPPID;

    const config = require('./config');
    assert.strictEqual(config.default.cliAadAppId, '31359c7f-bd7e-475c-86db-fdb8c937548e');
  });
});