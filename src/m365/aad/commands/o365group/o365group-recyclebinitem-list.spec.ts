import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./o365group-recyclebinitem-list');

describe(commands.O365GROUP_RECYCLEBINITEM_LIST, () => {
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
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_RECYCLEBINITEM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'mailNickname']);
  });

  it('lists deleted Microsoft 365 Groups in the tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return Promise.resolve({
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 2",
              "displayName": "Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: {} }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "deletedDateTime": null,
            "classification": null,
            "createdDateTime": "2017-12-07T13:58:01Z",
            "description": "Team 1",
            "displayName": "Team 1",
            "groupTypes": [
              "Unified"
            ],
            "mail": "team_1@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "team_1",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:team_1@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-07T13:58:01Z",
            "securityEnabled": false,
            "visibility": "Private"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "deletedDateTime": null,
            "classification": null,
            "createdDateTime": "2017-12-17T13:30:42Z",
            "description": "Team 2",
            "displayName": "Team 2",
            "groupTypes": [
              "Unified"
            ],
            "mail": "team_2@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "team_2",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:team_2@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-17T13:30:42Z",
            "securityEnabled": false,
            "visibility": "Private"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Deleted Microsoft 365 Groups in the tenant (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return Promise.resolve({
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": "2018-03-06T01:42:50Z",
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Deleted Team 1",
              "displayName": "Deleted Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "d_team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "d_team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:d_team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": "2018-03-06T01:42:50Z",
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Deleted Team 2",
              "displayName": "Deleted Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "d_team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "d_team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:d_team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
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
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "deletedDateTime": "2018-03-06T01:42:50Z",
            "classification": null,
            "createdDateTime": "2017-12-07T13:58:01Z",
            "description": "Deleted Team 1",
            "displayName": "Deleted Team 1",
            "groupTypes": [
              "Unified"
            ],
            "mail": "d_team_1@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "d_team_1",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:d_team_1@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-07T13:58:01Z",
            "securityEnabled": false,
            "visibility": "Private"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "deletedDateTime": "2018-03-06T01:42:50Z",
            "classification": null,
            "createdDateTime": "2017-12-17T13:30:42Z",
            "description": "Deleted Team 2",
            "displayName": "Deleted Team 2",
            "groupTypes": [
              "Unified"
            ],
            "mail": "d_team_2@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "d_team_2",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:d_team_2@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-17T13:30:42Z",
            "securityEnabled": false,
            "visibility": "Private"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Deleted Microsoft 365 Groups filtering on displayName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Deleted')&$top=100`) {
        return Promise.resolve({
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": "2018-03-06T01:42:50Z",
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Deleted Team 1",
              "displayName": "Deleted Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "d_team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "d_team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:d_team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": "2018-03-06T01:42:50Z",
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Deleted Team 2",
              "displayName": "Deleted Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "d_team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "d_team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:d_team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { displayName: 'Deleted' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "deletedDateTime": "2018-03-06T01:42:50Z",
            "classification": null,
            "createdDateTime": "2017-12-07T13:58:01Z",
            "description": "Deleted Team 1",
            "displayName": "Deleted Team 1",
            "groupTypes": [
              "Unified"
            ],
            "mail": "d_team_1@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "d_team_1",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:d_team_1@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-07T13:58:01Z",
            "securityEnabled": false,
            "visibility": "Private"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "deletedDateTime": "2018-03-06T01:42:50Z",
            "classification": null,
            "createdDateTime": "2017-12-17T13:30:42Z",
            "description": "Deleted Team 2",
            "displayName": "Deleted Team 2",
            "groupTypes": [
              "Unified"
            ],
            "mail": "d_team_2@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "d_team_2",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:d_team_2@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-17T13:30:42Z",
            "securityEnabled": false,
            "visibility": "Private"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Deleted Microsoft 365 Groups filtering on mailNickname', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'d_team')&$top=100`) {
        return Promise.resolve({
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": "2018-03-06T01:42:50Z",
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Deleted Team 1",
              "displayName": "Deleted Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "d_team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "d_team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:d_team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": "2018-03-06T01:42:50Z",
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Deleted Team 2",
              "displayName": "Deleted Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "d_team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "d_team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:d_team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { mailNickname: 'd_team' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "deletedDateTime": "2018-03-06T01:42:50Z",
            "classification": null,
            "createdDateTime": "2017-12-07T13:58:01Z",
            "description": "Deleted Team 1",
            "displayName": "Deleted Team 1",
            "groupTypes": [
              "Unified"
            ],
            "mail": "d_team_1@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "d_team_1",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:d_team_1@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-07T13:58:01Z",
            "securityEnabled": false,
            "visibility": "Private"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "deletedDateTime": "2018-03-06T01:42:50Z",
            "classification": null,
            "createdDateTime": "2017-12-17T13:30:42Z",
            "description": "Deleted Team 2",
            "displayName": "Deleted Team 2",
            "groupTypes": [
              "Unified"
            ],
            "mail": "d_team_2@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "d_team_2",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:d_team_2@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-17T13:30:42Z",
            "securityEnabled": false,
            "visibility": "Private"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Deleted Microsoft 365 Groups filtering on displayName and mailNickname', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Deleted') and startswith(MailNickname,'d_team')&$top=100`) {
        return Promise.resolve({
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": "2018-03-06T01:42:50Z",
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Deleted Team 1",
              "displayName": "Deleted Team 1",
              "groupTypes": [
                "Unified"
              ],
              "mail": "d_team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "d_team_1",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:d_team_1@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-07T13:58:01Z",
              "securityEnabled": false,
              "visibility": "Private"
            },
            {
              "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": "2018-03-06T01:42:50Z",
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Deleted Team 2",
              "displayName": "Deleted Team 2",
              "groupTypes": [
                "Unified"
              ],
              "mail": "d_team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "d_team_2",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:d_team_2@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-12-17T13:30:42Z",
              "securityEnabled": false,
              "visibility": "Private"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { displayName: 'Deleted', mailNickname: 'd_team' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "deletedDateTime": "2018-03-06T01:42:50Z",
            "classification": null,
            "createdDateTime": "2017-12-07T13:58:01Z",
            "description": "Deleted Team 1",
            "displayName": "Deleted Team 1",
            "groupTypes": [
              "Unified"
            ],
            "mail": "d_team_1@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "d_team_1",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:d_team_1@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-07T13:58:01Z",
            "securityEnabled": false,
            "visibility": "Private"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "deletedDateTime": "2018-03-06T01:42:50Z",
            "classification": null,
            "createdDateTime": "2017-12-17T13:30:42Z",
            "description": "Deleted Team 2",
            "displayName": "Deleted Team 2",
            "groupTypes": [
              "Unified"
            ],
            "mail": "d_team_2@contoso.onmicrosoft.com",
            "mailEnabled": true,
            "mailNickname": "d_team_2",
            "onPremisesLastSyncDateTime": null,
            "onPremisesProvisioningErrors": [],
            "onPremisesSecurityIdentifier": null,
            "onPremisesSyncEnabled": null,
            "preferredDataLocation": null,
            "proxyAddresses": [
              "SMTP:d_team_2@contoso.onmicrosoft.com"
            ],
            "renewedDateTime": "2017-12-17T13:30:42Z",
            "securityEnabled": false,
            "visibility": "Private"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying displayName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--displayName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying mailNickname', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--mailNickname') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

});