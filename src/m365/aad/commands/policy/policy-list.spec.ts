import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
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
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
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

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'isOrganizationDefault']);
  });

  it('retrieves the specified policy', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/policies/authorizationPolicy`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authorizationPolicy/$entity",
          "@odata.id": "https://graph.microsoft.com/v2/b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923/authorizationPolicy/authorizationPolicy",
          "id": "authorizationPolicy",
          "allowInvitesFrom": "everyone",
          "allowedToSignUpEmailBasedSubscriptions": true,
          "allowedToUseSSPR": true,
          "allowEmailVerifiedUsersToJoinOrganization": true,
          "blockMsolPowerShell": null,
          "displayName": "Authorization Policy",
          "description": "Used to manage authorization related settings across the company.",
          "defaultUserRolePermissions": {
            "allowedToCreateApps": true,
            "allowedToCreateSecurityGroups": true,
            "allowedToReadOtherUsers": true,
            "permissionGrantPoliciesAssigned": [
              "ManagePermissionGrantsForSelf.microsoft-user-default-legacy"
            ]
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        policyType: "authorization"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authorizationPolicy/$entity",
          "@odata.id": "https://graph.microsoft.com/v2/b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923/authorizationPolicy/authorizationPolicy",
          "id": "authorizationPolicy",
          "allowInvitesFrom": "everyone",
          "allowedToSignUpEmailBasedSubscriptions": true,
          "allowedToUseSSPR": true,
          "allowEmailVerifiedUsersToJoinOrganization": true,
          "blockMsolPowerShell": null,
          "displayName": "Authorization Policy",
          "description": "Used to manage authorization related settings across the company.",
          "defaultUserRolePermissions": {
            "allowedToCreateApps": true,
            "allowedToCreateSecurityGroups": true,
            "allowedToReadOtherUsers": true,
            "permissionGrantPoliciesAssigned": [
              "ManagePermissionGrantsForSelf.microsoft-user-default-legacy"
            ]
          }
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('retrieves all policies', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/policies/activityBasedTimeoutPolicies`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/activityBasedTimeoutPolicies",
          "value": []
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/authorizationPolicy`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authorizationPolicy/$entity",
          "@odata.id": "https://graph.microsoft.com/v2/b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923/authorizationPolicy/authorizationPolicy",
          "id": "authorizationPolicy",
          "allowInvitesFrom": "everyone",
          "allowedToSignUpEmailBasedSubscriptions": true,
          "allowedToUseSSPR": true,
          "allowEmailVerifiedUsersToJoinOrganization": true,
          "blockMsolPowerShell": null,
          "displayName": "Authorization Policy",
          "description": "Used to manage authorization related settings across the company.",
          "defaultUserRolePermissions": {
            "allowedToCreateApps": true,
            "allowedToCreateSecurityGroups": true,
            "allowedToReadOtherUsers": true,
            "permissionGrantPoliciesAssigned": [
              "ManagePermissionGrantsForSelf.microsoft-user-default-legacy"
            ]
          }
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

      if (opts.url === `https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/identitySecurityDefaultsEnforcementPolicy/$entity",
          "id": "00000000-0000-0000-0000-000000000005",
          "displayName": "Security Defaults",
          "description": "Security defaults is a set of basic identity security mechanisms recommended by Microsoft. When enabled, these recommendations will be automatically enforced in your organization. Administrators and users will be better protected from common identity related attacks.",
          "isEnabled": false
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
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/authorizationPolicy/$entity",
            "@odata.id": "https://graph.microsoft.com/v2/b30f2eac-f6b4-4f87-9dcb-cdf7ae1f8923/authorizationPolicy/authorizationPolicy",
            "id": "authorizationPolicy",
            "allowInvitesFrom": "everyone",
            "allowedToSignUpEmailBasedSubscriptions": true,
            "allowedToUseSSPR": true,
            "allowEmailVerifiedUsersToJoinOrganization": true,
            "blockMsolPowerShell": null,
            "displayName": "Authorization Policy",
            "description": "Used to manage authorization related settings across the company.",
            "defaultUserRolePermissions": {
              "allowedToCreateApps": true,
              "allowedToCreateSecurityGroups": true,
              "allowedToReadOtherUsers": true,
              "permissionGrantPoliciesAssigned": [
                "ManagePermissionGrantsForSelf.microsoft-user-default-legacy"
              ]
            }
          },
          {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#policies/identitySecurityDefaultsEnforcementPolicy/$entity",
            "id": "00000000-0000-0000-0000-000000000005",
            "displayName": "Security Defaults",
            "description": "Security defaults is a set of basic identity security mechanisms recommended by Microsoft. When enabled, these recommendations will be automatically enforced in your organization. Administrators and users will be better protected from common identity related attacks.",
            "isEnabled": false
          },
          {
            id: '457c8ef6-7a9c-4c9c-ba05-a12b7654c95a',
            deletedDateTime: null,
            definition: [
              '{ "TokenIssuancePolicy":{"TokenResponseSigningPolicy":"TokenOnly","SamlTokenVersion":"1.1","SigningAlgorithm":"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256","Version":1}}'
            ],
            displayName: 'TokenIssuancePolicy1',
            isOrganizationDefault: true
          },
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

  it('correctly handles API OData error for specified policies', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject("An error has occurred.");
    });

    command.action(logger, { options: { debug: false, policyType: "foo" } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error for all policies', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject("An error has occurred.");
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
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

  it('accepts policyType to be authorization', () => {
    const actual = command.validate({
      options:
      {
        policyType: "authorization"
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

  it('accepts policyType to be identitySecurityDefaultsEnforcement', () => {
    const actual = command.validate({
      options:
      {
        policyType: "identitySecurityDefaultsEnforcement"
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
    assert.notStrictEqual(actual, true);
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