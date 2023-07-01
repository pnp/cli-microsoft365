import assert from 'assert';

describe('Config', () => {
  it('returns process.env CLIMICROSOFT365_TENANT value', async () => {
    process.env.CLIMICROSOFT365_TENANT = 'tenant123';

    const config = await import(`./config.js#${Math.random()}`);
    assert.strictEqual(config.default.tenant, 'tenant123');
  });

  it('returns process.env CLIMICROSOFT365_AADAPPID value', async () => {
    process.env.CLIMICROSOFT365_AADAPPID = 'appId123';

    const config = await import(`./config.js#${Math.random()}`);
    assert.strictEqual(config.default.cliAadAppId, 'appId123');
  });

  it('returns default value since env CLIMICROSOFT365_TENANT not present', async () => {
    delete process.env.CLIMICROSOFT365_TENANT;

    const config = await import(`./config.js#${Math.random()}`);
    assert.strictEqual(config.default.tenant, 'common');
  });

  it('returns default value since env CLIMICROSOFT365_AADAPPID not present', async () => {
    delete process.env.CLIMICROSOFT365_AADAPPID;

    const config = await import(`./config.js#${Math.random()}`);
    assert.strictEqual(config.default.cliAadAppId, '31359c7f-bd7e-475c-86db-fdb8c937548e');
  });
});