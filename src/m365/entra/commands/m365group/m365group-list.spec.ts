import { Group } from '@microsoft/microsoft-graph-types';
import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './m365group-list.js';

describe(commands.M365GROUP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    assert.strictEqual(command.name, commands.M365GROUP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'mailNickname', 'siteUrl']);
  });

  it('lists Microsoft 365 Groups in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$expand=owners&$top=100`) {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ orphaned: true }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$expand=owners&$top=100`) {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, orphaned: true }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Team')&$top=100`) {
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

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'Team' }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'team')&$top=100`) {
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

    await command.action(logger, { options: commandOptionsSchema.parse({ mailNickname: 'team' }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'Team') and startswith(MailNickname,'team')&$top=100`) {
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

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'Team', mailNickname: 'team' }) });
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

  it('escapes special characters in the displayName filter', async () => {
    const displayName = 'Team\'s #';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(DisplayName,'${formatting.encodeQueryParameter(displayName)}')&$top=100`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: displayName }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and startswith(MailNickname,'${formatting.encodeQueryParameter(mailNickName)}')&$top=100`) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ mailNickname: mailNickName }) });
    assert(loggerLogSpy.calledWith([]));
  });

  it('lists Microsoft 365 Groups in the tenant served in pages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
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
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
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
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100&$skiptoken=X%2744537074090001000000000000000014000000C233BFA08475B84E8BF8C40335F8944D01000000000000000000000000000017312E322E3834302E3131333535362E312E342E32333331020000000000017D06501DC4C194438D57CFE494F81C1E%27`) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError('An error has occurred'));
  });

  it('lists all properties for output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

    await command.action(logger, { options: commandOptionsSchema.parse({ output: 'json' }) });
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

  it(`correctly shows deprecation warning for option 'includeSiteUrl'`, async () => {
    const chalk = (await import('chalk')).default;
    const loggerErrSpy = sinon.spy(logger, 'logToStderr');

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ includeSiteUrl: true }) });
    assert(loggerErrSpy.calledWith(chalk.yellow(`Parameter 'includeSiteUrl' is deprecated. Please use 'withSiteUrl' instead`)));

    sinonUtil.restore(loggerErrSpy);
  });

  it('include site URLs of Microsoft 365 Groups', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ withSiteUrl: true }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return Promise.resolve(<Group>{
          webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, withSiteUrl: true }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
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
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_1/Shared%20Documents" };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return { webUrl: "" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ withSiteUrl: true }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$top=100`) {
        return {
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
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/010d2f0a-0c17-4ec8-b694-e85bbe607013/drive?$select=webUrl`) {
        throw 'An error has occurred';
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0157132c-bf82-48ff-99e4-b19a74950fe0/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/team_2/Shared%20Documents" };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ withSiteUrl: true }) }), new CommandError('An error has occurred'));
  });

  it('passes validation if only includeSiteUrl option set', () => {
    const actual = commandOptionsSchema.safeParse({ includeSiteUrl: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if only withSiteUrl option set', () => {
    const actual = commandOptionsSchema.safeParse({ withSiteUrl: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if only orphaned option set', () => {
    const actual = commandOptionsSchema.safeParse({ orphaned: true });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if no options set', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });
});
