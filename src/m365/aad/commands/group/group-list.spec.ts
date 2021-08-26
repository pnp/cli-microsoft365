import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./group-list');

describe(commands.GROUP_LIST,()=>{
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(()=>{
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    assert.strictEqual(command.name.startsWith(commands.GROUP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'mailNickname','groupTypes', 'securityEnabled','mailEnabled','visibility']);
  });

  it('lists aad Groups in the tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups`) {
        return Promise.resolve({
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2021-03-13T14:04:39Z",
              "creationOptions": [
                "ProvisionGroupHomepage",
                "HubSiteId:00000000-0000-0000-0000-000000000000",
                "SPSiteLanguage:1033"
              ],
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "membershipRule": null,
              "membershipRuleProcessingState": null,
              "onPremisesDomainName": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesNetBiosName": null,
              "onPremisesSamAccountName": null,
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "preferredLanguage": null,
              "proxyAddresses": [
                "SMTP:CodeChallenge@dev1802.onmicrosoft.com"
              ],
              "renewedDateTime": "2021-03-13T14:04:39Z",
              "resourceBehaviorOptions": [],
              "resourceProvisioningOptions": [],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-14818455-1270970368-10757248-1844749995",
              "theme": null,
              "visibility": "Private",
              "onPremisesProvisioningErrors": []
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
            "deletedDateTime": null,
            "classification": null,
            "createdDateTime": "2021-03-13T14:04:39Z",
            "creationOptions": [
              "ProvisionGroupHomepage",
              "HubSiteId:00000000-0000-0000-0000-000000000000",
              "SPSiteLanguage:1033"
            ],
            "description": "Code Challenge",
            "displayName": "Code Challenge",
            "expirationDateTime": null,
            "groupTypes": [
              "Unified"
            ],
            "isAssignableToRole": null,
            "mail": "CodeChallenge@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "CodeChallenge",
            "membershipRule": null,
            "membershipRuleProcessingState": null,
            "onPremisesDomainName": null,
            "onPremisesLastSyncDateTime": null,
            "onPremisesNetBiosName": null,
            "onPremisesSamAccountName": null,
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "preferredLanguage": null,
            "proxyAddresses": [
              "SMTP:CodeChallenge@dev1802.onmicrosoft.com"
            ],
            "renewedDateTime": "2021-03-13T14:04:39Z",
            "resourceBehaviorOptions": [],
            "resourceProvisioningOptions": [],
            "securityEnabled": false,
            "securityIdentifier": "S-1-12-1-14818455-1270970368-10757248-1844749995",
            "theme": null,
            "visibility": "Private",
            "onPremisesProvisioningErrors": []
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists aad Groups in the tenant (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups`) {
        return Promise.resolve({
          "value": [
            {
              "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2021-03-13T14:04:39Z",
              "creationOptions": [
                "ProvisionGroupHomepage",
                "HubSiteId:00000000-0000-0000-0000-000000000000",
                "SPSiteLanguage:1033"
              ],
              "description": "Code Challenge",
              "displayName": "Code Challenge",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "CodeChallenge@dev1802.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "CodeChallenge",
              "membershipRule": null,
              "membershipRuleProcessingState": null,
              "onPremisesDomainName": null,
              "onPremisesLastSyncDateTime": null,
              "onPremisesNetBiosName": null,
              "onPremisesSamAccountName": null,
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "preferredLanguage": null,
              "proxyAddresses": [
                "SMTP:CodeChallenge@dev1802.onmicrosoft.com"
              ],
              "renewedDateTime": "2021-03-13T14:04:39Z",
              "resourceBehaviorOptions": [],
              "resourceProvisioningOptions": [],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-14818455-1270970368-10757248-1844749995",
              "theme": null,
              "visibility": "Private",
              "onPremisesProvisioningErrors": []
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { verbose: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "00e21c97-7800-4bc1-8024-a400aba6f46d",
            "deletedDateTime": null,
            "classification": null,
            "createdDateTime": "2021-03-13T14:04:39Z",
            "creationOptions": [
              "ProvisionGroupHomepage",
              "HubSiteId:00000000-0000-0000-0000-000000000000",
              "SPSiteLanguage:1033"
            ],
            "description": "Code Challenge",
            "displayName": "Code Challenge",
            "expirationDateTime": null,
            "groupTypes": [
              "Unified"
            ],
            "isAssignableToRole": null,
            "mail": "CodeChallenge@dev1802.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "CodeChallenge",
            "membershipRule": null,
            "membershipRuleProcessingState": null,
            "onPremisesDomainName": null,
            "onPremisesLastSyncDateTime": null,
            "onPremisesNetBiosName": null,
            "onPremisesSamAccountName": null,
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "preferredLanguage": null,
            "proxyAddresses": [
              "SMTP:CodeChallenge@dev1802.onmicrosoft.com"
            ],
            "renewedDateTime": "2021-03-13T14:04:39Z",
            "resourceBehaviorOptions": [],
            "resourceProvisioningOptions": [],
            "securityEnabled": false,
            "securityIdentifier": "S-1-12-1-14818455-1270970368-10757248-1844749995",
            "theme": null,
            "visibility": "Private",
            "onPremisesProvisioningErrors": []
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

