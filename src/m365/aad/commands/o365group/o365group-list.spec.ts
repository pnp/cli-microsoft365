import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./o365group-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.O365GROUP_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists Microsoft 365 Groups in the tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups in the tenant (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups without owners in the tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$expand=owners&$top=100`) {
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
              "visibility": "Private",
              "owners": []
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
              "visibility": "Private",
              "owners": [{
                "@odata.type": "#microsoft.graph.user",
                "id": "7343a4e9-159e-4736-a39d-f4ee2b2e1ff3",
                "displayName": "Joseph Velliah"
              },
              {
                "@odata.type": "#microsoft.graph.user",
                "id": "7343a4e9-159e-4736-a39d-f4ee2b2e1ff4",
                "displayName": "Bose Velliah"
              }]
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { orphaned: true } }, () => {
      try {
        assert([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups without owners in the tenant (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$expand=owners&$top=100`) {
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
              "visibility": "Private",
              "owners": []
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
              "visibility": "Private",
              "owners": [{
                "@odata.type": "#microsoft.graph.user",
                "id": "7343a4e9-159e-4736-a39d-f4ee2b2e1ff3",
                "displayName": "Joseph Velliah"
              },
              {
                "@odata.type": "#microsoft.graph.user",
                "id": "7343a4e9-159e-4736-a39d-f4ee2b2e1ff4",
                "displayName": "Bose Velliah"
              }]
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, orphaned: true } }, () => {
      try {
        assert([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups filtering on displayName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Team')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'Team' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups filtering on mailNickname', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'team')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, mailNickname: 'team' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups filtering on displayName and mailNickname', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Team') and startswith(MailNickname,'team')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'Team', mailNickname: 'team' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups in the tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups in the tenant (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups filtering on displayName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Team')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'Team' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups filtering on mailNickname', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'team')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, mailNickname: 'team' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups filtering on displayName and mailNickname', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Team') and startswith(MailNickname,'team')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'Team', mailNickname: 'team' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists deleted Microsoft 365 Groups in the tenant', (done) => {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, deleted: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Deleted Team 1",
            "mailNickname": "d_team_1",
            "deletedDateTime": "2018-03-06T01:42:50Z"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Deleted Team 2",
            "mailNickname": "d_team_2",
            "deletedDateTime": "2018-03-06T01:42:50Z"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Deleted Microsoft 365 Groups in the tenant (debug)', (done) => {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, deleted: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Deleted Team 1",
            "mailNickname": "d_team_1",
            "deletedDateTime": "2018-03-06T01:42:50Z"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Deleted Team 2",
            "mailNickname": "d_team_2",
            "deletedDateTime": "2018-03-06T01:42:50Z"
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { verbose: true, deleted: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Deleted Team 1",
            "mailNickname": "d_team_1",
            "deletedDateTime": "2018-03-06T01:42:50Z"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Deleted Team 2",
            "mailNickname": "d_team_2",
            "deletedDateTime": "2018-03-06T01:42:50Z"
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, deleted: true, displayName: 'Deleted' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Deleted Team 1",
            "mailNickname": "d_team_1",
            "deletedDateTime": "2018-03-06T01:42:50Z"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Deleted Team 2",
            "mailNickname": "d_team_2",
            "deletedDateTime": "2018-03-06T01:42:50Z"
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, deleted: true, mailNickname: 'd_team' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Deleted Team 1",
            "mailNickname": "d_team_1",
            "deletedDateTime": "2018-03-06T01:42:50Z"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Deleted Team 2",
            "mailNickname": "d_team_2",
            "deletedDateTime": "2018-03-06T01:42:50Z"
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, deleted: true, displayName: 'Deleted', mailNickname: 'd_team' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Deleted Team 1",
            "mailNickname": "d_team_1",
            "deletedDateTime": "2018-03-06T01:42:50Z"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Deleted Team 2",
            "mailNickname": "d_team_2",
            "deletedDateTime": "2018-03-06T01:42:50Z"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes special characters in the displayName filter', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Team''s%20%23')&$top=100`) {
        return Promise.resolve({
          "value": [
            {
              "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 1",
              "displayName": "Team's #1",
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
              "displayName": "Team's #2",
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, displayName: 'Team\'s #' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team's #1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team's #2",
            "mailNickname": "team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes special characters in the mailNickname filter', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'team''s%20%23')&$top=100`) {
        return Promise.resolve({
          "value": []
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, mailNickname: 'team\'s #' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft 365 Groups in the tenant served in pages', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27",
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27`) {
        return Promise.resolve({
          "value": [
            {
              "id": "310d2f0a-0c17-4ec8-b694-e85bbe607013",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-07T13:58:01Z",
              "description": "Team 3",
              "displayName": "Team 3",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_1@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_3",
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
              "id": "4157132c-bf82-48ff-99e4-b19a74950fe0",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-12-17T13:30:42Z",
              "description": "Team 4",
              "displayName": "Team 4",
              "groupTypes": [
                "Unified"
              ],
              "mail": "team_2@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "team_4",
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2"
          },
          {
            "id": "310d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 3",
            "mailNickname": "team_3"
          },
          {
            "id": "4157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 4",
            "mailNickname": "team_4"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving second page of Microsoft 365 Groups', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return Promise.resolve({
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27",
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27`) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists all properties for output json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
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

  it('include site URLs of Microsoft 365 Groups', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, includeSiteUrl: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1",
            "siteUrl": "https://contoso.sharepoint.com/sites/team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2",
            "siteUrl": "https://contoso.sharepoint.com/sites/team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('include site URLs of Microsoft 365 Groups (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, includeSiteUrl: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1",
            "siteUrl": "https://contoso.sharepoint.com/sites/team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2",
            "siteUrl": "https://contoso.sharepoint.com/sites/team_2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('include site URLs of Microsoft 365 Groups. one group without site', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: ""
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, includeSiteUrl: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "010d2f0a-0c17-4ec8-b694-e85bbe607013",
            "displayName": "Team 1",
            "mailNickname": "team_1",
            "siteUrl": "https://contoso.sharepoint.com/sites/team_1"
          },
          {
            "id": "0157132c-bf82-48ff-99e4-b19a74950fe0",
            "displayName": "Team 2",
            "mailNickname": "team_2",
            "siteUrl": ""
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving Microsoft 365 Group url', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return Promise.reject('An error has occurred');
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return Promise.resolve({
          webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, includeSiteUrl: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if both deleted and includeSiteUrl options set', () => {
    const actual = (command.validate() as CommandValidate)({ options: { deleted: true, includeSiteUrl: true } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if only deleted option set', () => {
    const actual = (command.validate() as CommandValidate)({ options: { deleted: true } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if only includeSiteUrl option set', () => {
    const actual = (command.validate() as CommandValidate)({ options: { includeSiteUrl: true } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if only orphaned option set', () => {
    const actual = (command.validate() as CommandValidate)({ options: { orphaned: true } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if no options set', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.strictEqual(actual, true);
  });
});