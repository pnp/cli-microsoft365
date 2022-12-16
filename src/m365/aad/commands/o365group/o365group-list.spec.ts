import { Group } from '@microsoft/microsoft-graph-types';
import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./o365group-list');

describe(commands.O365GROUP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'mailNickname', 'deletedDateTime', 'siteUrl']);
  });

  it('lists Microsoft 365 Groups in the tenant', async () => {
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

    await command.action(logger, { options: {} });
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
  });

  it('lists Microsoft 365 Groups in the tenant (debug)', async () => {
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

    await command.action(logger, { options: { debug: true } });
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
  });

  it('lists Microsoft 365 Groups without owners in the tenant', async () => {
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

    await command.action(logger, { options: { orphaned: true } });
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
  });

  it('lists Microsoft 365 Groups without owners in the tenant (debug)', async () => {
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

    await command.action(logger, { options: { debug: true, orphaned: true } });
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
  });

  it('lists Microsoft 365 Groups filtering on displayName', async () => {
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

    await command.action(logger, { options: { displayName: 'Team' } });
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
  });

  it('lists Microsoft 365 Groups filtering on mailNickname', async () => {
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

    await command.action(logger, { options: { mailNickname: 'team' } });
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
  });

  it('lists Microsoft 365 Groups filtering on displayName and mailNickname', async () => {
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

    await command.action(logger, { options: { displayName: 'Team', mailNickname: 'team' } });
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
  });

  it('lists deleted Microsoft 365 Groups in the tenant', async () => {
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

    await command.action(logger, { options: { deleted: true } });
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
  });

  it('lists Deleted Microsoft 365 Groups in the tenant (debug)', async () => {
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

    await command.action(logger, { options: { debug: true, deleted: true } });
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
  });

  it('lists Deleted Microsoft 365 Groups in the tenant (verbose)', async () => {
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

    await command.action(logger, { options: { verbose: true, deleted: true } });
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
  });

  it('lists Deleted Microsoft 365 Groups filtering on displayName', async () => {
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

    await command.action(logger, { options: { deleted: true, displayName: 'Deleted' } });
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
  });

  it('lists Deleted Microsoft 365 Groups filtering on mailNickname', async () => {
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

    await command.action(logger, { options: { deleted: true, mailNickname: 'd_team' } });
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
  });

  it('lists Deleted Microsoft 365 Groups filtering on displayName and mailNickname', async () => {
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

    await command.action(logger, { options: { deleted: true, displayName: 'Deleted', mailNickname: 'd_team' } });
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
  });

  it('escapes special characters in the displayName filter', async () => {
    const displayName = 'Team\'s #';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'${formatting.encodeQueryParameter(displayName)}')&$top=100`) {
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

    await command.action(logger, { options: { displayName: displayName } });
    assert(loggerLogSpy.calledWith([
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
    ]));
  });

  it('escapes special characters in the mailNickname filter', async () => {
    const mailNickName = 'team\'s #';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'${formatting.encodeQueryParameter(mailNickName)}')&$top=100`) {
        return Promise.resolve({
          "value": []
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { mailNickname: mailNickName } });
    assert(loggerLogSpy.calledWith([]));
  });

  it('lists Microsoft 365 Groups in the tenant served in pages', async () => {
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

    await command.action(logger, { options: {} });
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
      },

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
    ]));
  });

  it('handles error when retrieving second page of Microsoft 365 Groups', async () => {
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

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred'));
  });

  it('lists all properties for output json', async () => {
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

    await command.action(logger, { options: { output: 'json' } });
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
  });

  it('include site URLs of Microsoft 365 Groups', async () => {
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
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { includeSiteUrl: true } });
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
        "visibility": "Private",
        "siteUrl": "https://contoso.sharepoint.com/sites/team_1"
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
        "siteUrl": "https://contoso.sharepoint.com/sites/team_2"
      }
    ]));
  });

  it('include site URLs of Microsoft 365 Groups (debug)', async () => {
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
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, includeSiteUrl: true } });
    assert(loggerLogSpy.calledWith([
      <Group>{
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
        "siteUrl": "https://contoso.sharepoint.com/sites/team_1"
      },
      <Group>{
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
        "siteUrl": "https://contoso.sharepoint.com/sites/team_2"
      }
    ]));
  });

  it('include site URLs of Microsoft 365 Groups. one group without site', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return Promise.resolve({
          "value": [
            <Group>{
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
            <Group>{
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
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return Promise.resolve(<Group>{
          webUrl: ""
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { includeSiteUrl: true } });
    assert(loggerLogSpy.calledWith([
      <Group>{
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
        "siteUrl": "https://contoso.sharepoint.com/sites/team_1"
      },
      <Group>{
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
        "siteUrl": ""
      }
    ]));
  });

  it('handles error when retrieving Microsoft 365 Group url', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return Promise.resolve({
          "value": [
            <Group>{
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
            <Group>{
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
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents"
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { includeSiteUrl: true } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if both deleted and includeSiteUrl options set', async () => {
    const actual = await command.validate({ options: { deleted: true, includeSiteUrl: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if only deleted option set', async () => {
    const actual = await command.validate({ options: { deleted: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only includeSiteUrl option set', async () => {
    const actual = await command.validate({ options: { includeSiteUrl: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only orphaned option set', async () => {
    const actual = await command.validate({ options: { orphaned: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if no options set', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
