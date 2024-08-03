import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './m365group-recyclebinitem-list.js';
import aadCommands from '../../aadCommands.js';

describe(commands.M365GROUP_RECYCLEBINITEM_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
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
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_RECYCLEBINITEM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.M365GROUP_RECYCLEBINITEM_LIST]);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'mailNickname']);
  });

  it('lists deleted Microsoft 365 Groups in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
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
        };
      }

      throw 'Invalid request';
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

  it('lists Deleted Microsoft 365 Groups in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Deleted')&$top=100`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: 'Deleted' } });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'d_team')&$top=100`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupMailNickname: 'd_team' } });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Deleted') and startswith(MailNickname,'d_team')&$top=100`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: 'Deleted', groupMailNickname: 'd_team' } });
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

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: { mailNickname: 'd_team' } }), new CommandError(errorMessage));
  });

  it('supports specifying groupName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--groupName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying groupMailNickname', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--groupMailNickname') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
