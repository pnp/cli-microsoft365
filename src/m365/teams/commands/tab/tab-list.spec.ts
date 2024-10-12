import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tab-list.js';

describe(commands.TAB_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    assert.strictEqual(command.name, commands.TAB_LIST);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'teamsAppTabId']);
  });

  it('fails validation for a incorrect channelId missing leading 19:.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for a incorrect channelId missing trailing @thread.skype.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles teams tabs request failure', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs?$expand=teamsApp`) {
        throw {
          "error": {
            "code": "InvalidRequest",
            "message": "Channel id is not in a valid format: 29:d09d9792d59544af846fa19c98b6acc6@thread.skype",
            "innerError": {
              "request-id": "d2ad1b93-0bf8-43b7-aba1-99775175fa33",
              "date": "2019-03-30T22:37:30"
            }
          }
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype'
      }
    } as any), new CommandError('Channel id is not in a valid format: 29:d09d9792d59544af846fa19c98b6acc6@thread.skype'));
  });

  it('correctly lists all tabs in a Microsoft Teams channel in text format', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs?$expand=teamsApp`) {
        return {
          value: [
            {
              "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4",
              "displayName": "Document%20Library",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3afbd8cf48-e450-463a-8636-231307dda5f6?label=Document%2520Library&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
              "configuration": { "entityId": null, "contentUrl": "https://contoso.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente", "removeUrl": null, "websiteUrl": null },
              "teamsApp": { "id": "com.microsoft.teamspace.tab.files.sharepoint", "externalId": null, "displayName": "Document Library", "distributionMethod": "store" }
            },
            {
              "id": "ba38f554-9ce6-4719-bc9b-e38e4ca16860",
              "displayName": "CLI-Microsoft365",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a5a2f05cf-0b6d-4012-b296-342d89158248?webUrl=https%3a%2f%2fgithub.com%2fpnp%2fcli-microsoft365&label=CLI-Microsoft365&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
              "configuration": { "entityId": null, "contentUrl": "https://github.com/pnp/cli-microsoft365", "removeUrl": null, "websiteUrl": "https://github.com/pnp/cli-microsoft365", "dateAdded": "2019-03-28T18:35:53.81Z" },
              "teamsApp": { "id": "com.microsoft.teamspace.tab.web", "externalId": null, "displayName": "Website", "distributionMethod": "store" }
            },
            {
              "id": "b6c511f1-3ad7-4111-8a82-36b13aad4c9e",
              "displayName": "Word",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a00b9c26c-b0f4-4c1a-98fa-426aded95ca3?label=Word&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
              "configuration": { "entityId": "C6DBBF49-0290-4194-B3DA-319A72014FD6", "contentUrl": "https://contoso.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente/General/Kopieren und Verschieben von Dateien in Office365.docx", "removeUrl": null, "websiteUrl": null },
              "teamsApp": { "id": "com.microsoft.teamspace.tab.file.staticviewer.word", "externalId": null, "displayName": "Word", "distributionMethod": "store" }
            },
            {
              "id": "d9e972d8-e93d-4b87-beb2-3698912398ea",
              "displayName": "Wiki",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a9d44e015-ae5c-47dc-9577-5af76609e2b0?label=Wiki&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
              "configuration": { "entityId": null, "contentUrl": null, "removeUrl": null, "websiteUrl": null, "wikiTabName": "Wiki", "wikiTabId": 1, "wikiDefaultTab": true },
              "teamsApp": { "id": "com.microsoft.teamspace.tab.wiki", "externalId": null, "displayName": "Wiki", "distributionMethod": "store" }
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        output: 'text'
      }
    });

    assert(loggerLogSpy.calledWith([{
      "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4",
      "displayName": "Document%20Library",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3afbd8cf48-e450-463a-8636-231307dda5f6?label=Document%2520Library&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
      "configuration": {
        "entityId": null,
        "contentUrl": "https://contoso.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente",
        "removeUrl": null,
        "websiteUrl": null
      },
      "teamsApp": {
        "id": "com.microsoft.teamspace.tab.files.sharepoint",
        "externalId": null,
        "displayName": "Document Library",
        "distributionMethod": "store"
      },
      "teamsAppTabId": "com.microsoft.teamspace.tab.files.sharepoint"
    },
    {
      "id": "ba38f554-9ce6-4719-bc9b-e38e4ca16860",
      "displayName": "CLI-Microsoft365",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a5a2f05cf-0b6d-4012-b296-342d89158248?webUrl=https%3a%2f%2fgithub.com%2fpnp%2fcli-microsoft365&label=CLI-Microsoft365&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
      "configuration": {
        "entityId": null,
        "contentUrl": "https://github.com/pnp/cli-microsoft365",
        "removeUrl": null,
        "websiteUrl": "https://github.com/pnp/cli-microsoft365",
        "dateAdded": "2019-03-28T18:35:53.81Z"
      },
      "teamsApp": {
        "id": "com.microsoft.teamspace.tab.web",
        "externalId": null,
        "displayName": "Website",
        "distributionMethod": "store"
      },
      "teamsAppTabId": "com.microsoft.teamspace.tab.web"
    },
    {
      "id": "b6c511f1-3ad7-4111-8a82-36b13aad4c9e",
      "displayName": "Word",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a00b9c26c-b0f4-4c1a-98fa-426aded95ca3?label=Word&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
      "configuration": {
        "entityId": "C6DBBF49-0290-4194-B3DA-319A72014FD6",
        "contentUrl": "https://contoso.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente/General/Kopieren und Verschieben von Dateien in Office365.docx",
        "removeUrl": null,
        "websiteUrl": null
      },
      "teamsApp": {
        "id": "com.microsoft.teamspace.tab.file.staticviewer.word",
        "externalId": null,
        "displayName": "Word",
        "distributionMethod": "store"
      },
      "teamsAppTabId": "com.microsoft.teamspace.tab.file.staticviewer.word"
    },
    {
      "id": "d9e972d8-e93d-4b87-beb2-3698912398ea",
      "displayName": "Wiki",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a9d44e015-ae5c-47dc-9577-5af76609e2b0?label=Wiki&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
      "configuration": {
        "entityId": null,
        "contentUrl": null,
        "removeUrl": null,
        "websiteUrl": null,
        "wikiTabName": "Wiki",
        "wikiTabId": 1,
        "wikiDefaultTab": true
      },
      "teamsApp": {
        "id": "com.microsoft.teamspace.tab.wiki",
        "externalId": null,
        "displayName": "Wiki",
        "distributionMethod": "store"
      },
      "teamsAppTabId": "com.microsoft.teamspace.tab.wiki"
    }]));
  });

  it('correctly lists all tabs in a Microsoft Teams channel in json format', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs?$expand=teamsApp`) {
        return {
          value: [
            {
              "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4",
              "displayName": "Document%20Library",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3afbd8cf48-e450-463a-8636-231307dda5f6?label=Document%2520Library&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
              "configuration": { "entityId": null, "contentUrl": "https://contoso.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente", "removeUrl": null, "websiteUrl": null },
              "teamsApp": { "id": "com.microsoft.teamspace.tab.files.sharepoint", "externalId": null, "displayName": "Document Library", "distributionMethod": "store" }
            },
            {
              "id": "ba38f554-9ce6-4719-bc9b-e38e4ca16860",
              "displayName": "CLI-Microsoft365",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a5a2f05cf-0b6d-4012-b296-342d89158248?webUrl=https%3a%2f%2fgithub.com%2fpnp%2fcli-microsoft365&label=CLI-Microsoft365&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
              "configuration": { "entityId": null, "contentUrl": "https://github.com/pnp/cli-microsoft365", "removeUrl": null, "websiteUrl": "https://github.com/pnp/cli-microsoft365", "dateAdded": "2019-03-28T18:35:53.81Z" },
              "teamsApp": { "id": "com.microsoft.teamspace.tab.web", "externalId": null, "displayName": "Website", "distributionMethod": "store" }
            },
            {
              "id": "b6c511f1-3ad7-4111-8a82-36b13aad4c9e",
              "displayName": "Word",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a00b9c26c-b0f4-4c1a-98fa-426aded95ca3?label=Word&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
              "configuration": { "entityId": "C6DBBF49-0290-4194-B3DA-319A72014FD6", "contentUrl": "https://contoso.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente/General/Kopieren und Verschieben von Dateien in Office365.docx", "removeUrl": null, "websiteUrl": null },
              "teamsApp": { "id": "com.microsoft.teamspace.tab.file.staticviewer.word", "externalId": null, "displayName": "Word", "distributionMethod": "store" }
            },
            {
              "id": "d9e972d8-e93d-4b87-beb2-3698912398ea",
              "displayName": "Wiki",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a9d44e015-ae5c-47dc-9577-5af76609e2b0?label=Wiki&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
              "configuration": { "entityId": null, "contentUrl": null, "removeUrl": null, "websiteUrl": null, "wikiTabName": "Wiki", "wikiTabId": 1, "wikiDefaultTab": true },
              "teamsApp": { "id": "com.microsoft.teamspace.tab.wiki", "externalId": null, "displayName": "Wiki", "distributionMethod": "store" }
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        output: 'json'
      }
    });

    assert(loggerLogSpy.calledWith([{
      "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4",
      "displayName": "Document%20Library",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3afbd8cf48-e450-463a-8636-231307dda5f6?label=Document%2520Library&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
      "configuration": {
        "entityId": null,
        "contentUrl": "https://contoso.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente",
        "removeUrl": null,
        "websiteUrl": null
      },
      "teamsApp": {
        "id": "com.microsoft.teamspace.tab.files.sharepoint",
        "externalId": null,
        "displayName": "Document Library",
        "distributionMethod": "store"
      }
    },
    {
      "id": "ba38f554-9ce6-4719-bc9b-e38e4ca16860",
      "displayName": "CLI-Microsoft365",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a5a2f05cf-0b6d-4012-b296-342d89158248?webUrl=https%3a%2f%2fgithub.com%2fpnp%2fcli-microsoft365&label=CLI-Microsoft365&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
      "configuration": {
        "entityId": null,
        "contentUrl": "https://github.com/pnp/cli-microsoft365",
        "removeUrl": null,
        "websiteUrl": "https://github.com/pnp/cli-microsoft365",
        "dateAdded": "2019-03-28T18:35:53.81Z"
      },
      "teamsApp": {
        "id": "com.microsoft.teamspace.tab.web",
        "externalId": null,
        "displayName": "Website",
        "distributionMethod": "store"
      }
    },
    {
      "id": "b6c511f1-3ad7-4111-8a82-36b13aad4c9e",
      "displayName": "Word",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a00b9c26c-b0f4-4c1a-98fa-426aded95ca3?label=Word&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
      "configuration": {
        "entityId": "C6DBBF49-0290-4194-B3DA-319A72014FD6",
        "contentUrl": "https://contoso.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente/General/Kopieren und Verschieben von Dateien in Office365.docx",
        "removeUrl": null,
        "websiteUrl": null
      },
      "teamsApp": {
        "id": "com.microsoft.teamspace.tab.file.staticviewer.word",
        "externalId": null,
        "displayName": "Word",
        "distributionMethod": "store"
      }
    },
    {
      "id": "d9e972d8-e93d-4b87-beb2-3698912398ea",
      "displayName": "Wiki",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a9d44e015-ae5c-47dc-9577-5af76609e2b0?label=Wiki&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
      "configuration": {
        "entityId": null,
        "contentUrl": null,
        "removeUrl": null,
        "websiteUrl": null,
        "wikiTabName": "Wiki",
        "wikiTabId": 1,
        "wikiDefaultTab": true
      },
      "teamsApp": {
        "id": "com.microsoft.teamspace.tab.wiki",
        "externalId": null,
        "displayName": "Wiki",
        "distributionMethod": "store"
      }
    }]));
  });
});