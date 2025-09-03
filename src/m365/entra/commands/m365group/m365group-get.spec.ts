import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './m365group-get.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

describe(commands.M365GROUP_GET, () => {
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
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(true);
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
    assert.strictEqual(command.name, commands.M365GROUP_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the Microsoft 365 Group specified by id', async () => {

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return {
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`) {
        return {
          "allowExternalSenders": false,
          "autoSubscribeNewMembers": false,
          "isSubscribedByMail": false,
          "hideFromOutlookClients": false,
          "hideFromAddressLists": false
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' }) });
    assert(loggerLogSpy.calledWith({
      "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2017-11-29T03:27:05Z",
      "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
      "displayName": "Finance",
      "groupTypes": [
        "Unified"
      ],
      "mail": "finance@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "finance",
      "onPremisesLastSyncDateTime": null,
      "onPremisesProvisioningErrors": [],
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "proxyAddresses": [
        "SMTP:finance@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2017-11-29T03:27:05Z",
      "securityEnabled": false,
      "visibility": "Public",
      "allowExternalSenders": false,
      "autoSubscribeNewMembers": false,
      "isSubscribedByMail": false,
      "hideFromOutlookClients": false,
      "hideFromAddressLists": false
    }));
  });

  it('retrieves information about the Microsoft 365 Group specified by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Finance'`) {
        return {
          "value": [
            {
              "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2017-11-29T03:27:05Z",
              "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
              "displayName": "Finance",
              "groupTypes": [
                "Unified"
              ],
              "mail": "finance@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "finance",
              "onPremisesLastSyncDateTime": null,
              "onPremisesProvisioningErrors": [],
              "onPremisesSecurityIdentifier": null,
              "onPremisesSyncEnabled": null,
              "preferredDataLocation": null,
              "proxyAddresses": [
                "SMTP:finance@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2017-11-29T03:27:05Z",
              "securityEnabled": false,
              "visibility": "Public"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`) {
        return {
          "allowExternalSenders": false,
          "autoSubscribeNewMembers": false,
          "isSubscribedByMail": false,
          "hideFromOutlookClients": false,
          "hideFromAddressLists": false
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'Finance' }) });
    assert(loggerLogSpy.calledWith({
      "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2017-11-29T03:27:05Z",
      "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
      "displayName": "Finance",
      "groupTypes": [
        "Unified"
      ],
      "mail": "finance@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "finance",
      "onPremisesLastSyncDateTime": null,
      "onPremisesProvisioningErrors": [],
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "proxyAddresses": [
        "SMTP:finance@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2017-11-29T03:27:05Z",
      "securityEnabled": false,
      "visibility": "Public",
      "allowExternalSenders": false,
      "autoSubscribeNewMembers": false,
      "isSubscribedByMail": false,
      "hideFromOutlookClients": false,
      "hideFromAddressLists": false
    }));
  });

  it('retrieves information about the specified Microsoft 365 Group (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return {
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`) {
        return {
          "allowExternalSenders": false,
          "autoSubscribeNewMembers": false,
          "isSubscribedByMail": false,
          "hideFromOutlookClients": false,
          "hideFromAddressLists": false
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' }) });
    assert(loggerLogSpy.calledWith({
      "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2017-11-29T03:27:05Z",
      "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
      "displayName": "Finance",
      "groupTypes": [
        "Unified"
      ],
      "mail": "finance@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "finance",
      "onPremisesLastSyncDateTime": null,
      "onPremisesProvisioningErrors": [],
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "proxyAddresses": [
        "SMTP:finance@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2017-11-29T03:27:05Z",
      "securityEnabled": false,
      "visibility": "Public",
      "allowExternalSenders": false,
      "autoSubscribeNewMembers": false,
      "isSubscribedByMail": false,
      "hideFromOutlookClients": false,
      "hideFromAddressLists": false
    }));
  });

  it(`correctly shows deprecation warning for option 'includeSiteUrl'`, async () => {
    const chalk = (await import('chalk')).default;
    const loggerErrSpy = sinon.spy(logger, 'logToStderr');

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return {
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`) {
        return {
          "allowExternalSenders": false,
          "autoSubscribeNewMembers": false,
          "isSubscribedByMail": false,
          "hideFromOutlookClients": false,
          "hideFromAddressLists": false
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/finance/Shared%20Documents" };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true }) });
    assert(loggerErrSpy.calledWith(chalk.yellow(`Parameter 'includeSiteUrl' is deprecated. Please use 'withSiteUrl' instead`)));

    sinonUtil.restore(loggerErrSpy);
  });

  it('retrieves information about the specified Microsoft 365 Group including its site URL', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return {
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`) {
        return {
          "allowExternalSenders": false,
          "autoSubscribeNewMembers": false,
          "isSubscribedByMail": false,
          "hideFromOutlookClients": false,
          "hideFromAddressLists": false
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/finance/Shared%20Documents" };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', withSiteUrl: true }) });
    assert(loggerLogSpy.calledWith({
      "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2017-11-29T03:27:05Z",
      "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
      "displayName": "Finance",
      "groupTypes": [
        "Unified"
      ],
      "mail": "finance@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "finance",
      "onPremisesLastSyncDateTime": null,
      "onPremisesProvisioningErrors": [],
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "proxyAddresses": [
        "SMTP:finance@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2017-11-29T03:27:05Z",
      "securityEnabled": false,
      "siteUrl": "https://contoso.sharepoint.com/sites/finance",
      "visibility": "Public",
      "allowExternalSenders": false,
      "autoSubscribeNewMembers": false,
      "isSubscribedByMail": false,
      "hideFromOutlookClients": false,
      "hideFromAddressLists": false
    }));
  });

  it('retrieves information about the specified Microsoft 365 Group including its site URL (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return {
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`) {
        return {
          "allowExternalSenders": false,
          "autoSubscribeNewMembers": false,
          "isSubscribedByMail": false,
          "hideFromOutlookClients": false,
          "hideFromAddressLists": false
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/finance/Shared%20Documents" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', withSiteUrl: true }) });
    assert(loggerLogSpy.calledWith({
      "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2017-11-29T03:27:05Z",
      "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
      "displayName": "Finance",
      "groupTypes": [
        "Unified"
      ],
      "mail": "finance@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "finance",
      "onPremisesLastSyncDateTime": null,
      "onPremisesProvisioningErrors": [],
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "proxyAddresses": [
        "SMTP:finance@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2017-11-29T03:27:05Z",
      "securityEnabled": false,
      "siteUrl": "https://contoso.sharepoint.com/sites/finance",
      "visibility": "Public",
      "allowExternalSenders": false,
      "autoSubscribeNewMembers": false,
      "isSubscribedByMail": false,
      "hideFromOutlookClients": false,
      "hideFromAddressLists": false
    }));
  });

  it('retrieves information about the specified Microsoft 365 Group including its site URL (group has no site)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {
        return {
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "deletedDateTime": null,
          "classification": null,
          "createdDateTime": "2017-11-29T03:27:05Z",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "displayName": "Finance",
          "groupTypes": [
            "Unified"
          ],
          "mail": "finance@contoso.onmicrosoft.com",
          "mailEnabled": true,
          "mailNickname": "finance",
          "onPremisesLastSyncDateTime": null,
          "onPremisesProvisioningErrors": [],
          "onPremisesSecurityIdentifier": null,
          "onPremisesSyncEnabled": null,
          "preferredDataLocation": null,
          "proxyAddresses": [
            "SMTP:finance@contoso.onmicrosoft.com"
          ],
          "renewedDateTime": "2017-11-29T03:27:05Z",
          "securityEnabled": false,
          "visibility": "Public"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`) {
        return {
          "allowExternalSenders": false,
          "autoSubscribeNewMembers": false,
          "isSubscribedByMail": false,
          "hideFromOutlookClients": false,
          "hideFromAddressLists": false
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return { webUrl: "" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', withSiteUrl: true }) });
    assert(loggerLogSpy.calledWith({
      "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2017-11-29T03:27:05Z",
      "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
      "displayName": "Finance",
      "groupTypes": [
        "Unified"
      ],
      "mail": "finance@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "finance",
      "onPremisesLastSyncDateTime": null,
      "onPremisesProvisioningErrors": [],
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "proxyAddresses": [
        "SMTP:finance@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2017-11-29T03:27:05Z",
      "securityEnabled": false,
      "visibility": "Public",
      "siteUrl": "",
      "allowExternalSenders": false,
      "autoSubscribeNewMembers": false,
      "isSubscribedByMail": false,
      "hideFromOutlookClients": false,
      "hideFromAddressLists": false
    }));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' }) }), new CommandError(errorMessage));
  });

  it('fails validation when id and displayName are not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when both id and displayName are specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
      displayName: 'Finance'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ id: '123' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the id is a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when displayName is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      displayName: 'Finance'
    });
    assert.strictEqual(actual.success, true);
  });

  it('shows error when the group is not a unified group', async () => {
    sinon.stub(entraGroup, 'getGroupById').resolves({
      "id": "3f04e370-cbc6-4091-80fe-1d038be2ad06",
      "deletedDateTime": null,
      "classification": null,
      "createdDateTime": "2017-11-29T03:27:05Z",
      "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
      "displayName": "Finance",
      "groupTypes": [],
      "mail": "finance@contoso.onmicrosoft.com",
      "mailEnabled": true,
      "mailNickname": "finance",
      "onPremisesLastSyncDateTime": null,
      "onPremisesProvisioningErrors": [],
      "onPremisesSecurityIdentifier": null,
      "onPremisesSyncEnabled": null,
      "preferredDataLocation": null,
      "proxyAddresses": [
        "SMTP:finance@contoso.onmicrosoft.com"
      ],
      "renewedDateTime": "2017-11-29T03:27:05Z",
      "securityEnabled": false,
      "visibility": "Public"
    });
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: groupId }) }),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });
});
