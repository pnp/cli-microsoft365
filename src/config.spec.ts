import * as assert from 'assert';

describe('Config', () => {
  before(() => {
    delete require.cache[require.resolve('./config')];
  });

  afterEach(() => {
    delete require.cache[require.resolve('./config')];
  });

  it('returns process.env OFFICE365CLI_TENANT value', () => {
    process.env.OFFICE365CLI_TENANT = 'tenant123';

    const config = require('./config');
    assert.equal(config.default.tenant, 'tenant123');
  });

  it('returns process.env OFFICE365CLI_AADAPPID value', () => {
    process.env.OFFICE365CLI_AADAPPID = 'appId123';

    const config = require('./config');
    assert.equal(config.default.cliAadAppId, 'appId123');
  });

  it('returns default value since env OFFICE365CLI_TENANT not present', () => {
    delete process.env.OFFICE365CLI_TENANT;

    const config = require('./config');
    assert.equal(config.default.tenant, 'common');
  });

  it('returns default value since env OFFICE365CLI_AADAPPID not present', () => {
    delete process.env.OFFICE365CLI_AADAPPID;

    const config = require('./config');
    assert.equal(config.default.cliAadAppId, '31359c7f-bd7e-475c-86db-fdb8c937548e');
  });
});