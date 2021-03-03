import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./policy-list');

describe(commands.POLICY_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.POLICY_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the specified policies', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/policies/tokenLifetimePolicies`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/tokenLifetimePolicies",
          "value": [
            {
              id: 'a457c42c-0f2e-4a25-be2a-545e840add1f',
              deletedDateTime: null,
              definition: [
                '{"TokenLifetimePolicy":{"Version":1,"AccessTokenLifetime":"8:00:00"}}'
              ],
              displayName: 'TokenLifetimePolicy1',
              isOrganizationDefault: true
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        policyType: "tokenLifetime"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            id: 'a457c42c-0f2e-4a25-be2a-545e840add1f',
            deletedDateTime: null,
            definition: [
              '{"TokenLifetimePolicy":{"Version":1,"AccessTokenLifetime":"8:00:00"}}'
            ],
            displayName: 'TokenLifetimePolicy1',
            isOrganizationDefault: true
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all the policies', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/policies/activityBasedTimeoutPolicies`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/activityBasedTimeoutPolicies",
          "value": []
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/claimsMappingPolicies`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/claimsMappingPolicies",
          "value": []
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/homeRealmDiscoveryPolicies`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/homeRealmDiscoveryPolicies",
          "value": []
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/tokenLifetimePolicies`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/tokenLifetimePolicies",
          "value": [
            {
              id: 'a457c42c-0f2e-4a25-be2a-545e840add1f',
              deletedDateTime: null,
              definition: [
                '{"TokenLifetimePolicy":{"Version":1,"AccessTokenLifetime":"8:00:00"}}'
              ],
              displayName: 'TokenLifetimePolicy1',
              isOrganizationDefault: true
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/tokenIssuancePolicies`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/tokenIssuancePolicies",
          "value": [
            {
              id: '457c8ef6-7a9c-4c9c-ba05-a12b7654c95a',
              deletedDateTime: null,
              definition: [
                '{ "TokenIssuancePolicy":{"TokenResponseSigningPolicy":"TokenOnly","SamlTokenVersion":"1.1","SigningAlgorithm":"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256","Version":1}}'
              ],
              displayName: 'TokenIssuancePolicy1',
              isOrganizationDefault: true
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            id: 'a457c42c-0f2e-4a25-be2a-545e840add1f',
            deletedDateTime: null,
            definition: [
              '{"TokenLifetimePolicy":{"Version":1,"AccessTokenLifetime":"8:00:00"}}'
            ],
            displayName: 'TokenLifetimePolicy1',
            isOrganizationDefault: true
          },
          {
            id: '457c8ef6-7a9c-4c9c-ba05-a12b7654c95a',
            deletedDateTime: null,
            definition: [
              '{ "TokenIssuancePolicy":{"TokenResponseSigningPolicy":"TokenOnly","SamlTokenVersion":"1.1","SigningAlgorithm":"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256","Version":1}}'
            ],
            displayName: 'TokenIssuancePolicy1',
            isOrganizationDefault: true
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error for specified policies', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject("Resource not found for the segment 'foo'.");
    });

    command.action(logger, { options: { debug: false, policyType: "foo" } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Resource not found for the segment 'foo'.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('accepts policyType to be activityBasedTimeout', () => {
    const actual = command.validate({
      options:
      {
        policyType: "activityBasedTimeout"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts policyType to be claimsMapping', () => {
    const actual = command.validate({
      options:
      {
        policyType: "claimsMapping"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts policyType to be homeRealmDiscovery', () => {
    const actual = command.validate({
      options:
      {
        policyType: "homeRealmDiscovery"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts policyType to be tokenLifetime', () => {
    const actual = command.validate({
      options:
      {
        policyType: "tokenLifetime"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts policyType to be tokenIssuance', () => {
    const actual = command.validate({
      options:
      {
        policyType: "tokenIssuance"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid policyType', () => {
    const policyType = 'foo';
    const actual = command.validate({
      options: {
        policyType: policyType
      }
    });
    assert.strictEqual(actual, `${policyType} is not a valid policyType. Allowed values are activityBasedTimeout|claimsMapping|homeRealmDiscovery|tokenLifetime|tokenIssuance`);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});