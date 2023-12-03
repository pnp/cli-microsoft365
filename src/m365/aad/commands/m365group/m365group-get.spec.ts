import assert from 'assert';
import sinon from 'sinon';
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
import { aadGroup } from '../../../../utils/aadGroup.js';

describe(commands.M365GROUP_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(true);
    auth.service.connected = true;
    commandInfo = cli.getCommandInfo(command);
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
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the specified Microsoft 365 Group', async () => {
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

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
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
      "visibility": "Public"
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

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
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
      "visibility": "Public"
    }));
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/finance/Shared%20Documents" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true } });
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
      "visibility": "Public"
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return { webUrl: "https://contoso.sharepoint.com/sites/finance/Shared%20Documents" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true } });
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
      "visibility": "Public"
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

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/1caf7dcd-7e83-4c3a-94f7-932a1299c844/drive?$select=webUrl`) {
        return { webUrl: "" };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', includeSiteUrl: true } });
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
      "siteUrl": ""
    }));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }), new CommandError(errorMessage));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('shows error when the group is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(aadGroup.isUnifiedGroup);
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { id: groupId } } as any),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });
});
