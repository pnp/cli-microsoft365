import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./tab-list');

describe(commands.TAB_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TAB_LIST), true);
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

  it('correctly handles teams tabs request failure', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs?$expand=teamsApp`) {
        return Promise.reject({
          "error": {
            "code": "InvalidRequest",
            "message": "Channel id is not in a valid format: 29:d09d9792d59544af846fa19c98b6acc6@thread.skype",
            "innerError": {
              "request-id": "d2ad1b93-0bf8-43b7-aba1-99775175fa33",
              "date": "2019-03-30T22:37:30"
            }
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype'
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError("Channel id is not in a valid format: 29:d09d9792d59544af846fa19c98b6acc6@thread.skype")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists all tabs in a Microsoft Teams channel', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs?$expand=teamsApp`) {
        return Promise.resolve({
          value: [
            {
              "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4",
              "displayName": "Document%20Library",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3afbd8cf48-e450-463a-8636-231307dda5f6?label=Document%2520Library&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
              "configuration": { "entityId": null, "contentUrl": "https://thomy.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente", "removeUrl": null, "websiteUrl": null },
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
              "configuration": { "entityId": "C6DBBF49-0290-4194-B3DA-319A72014FD6", "contentUrl": "https://thomy.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente/General/Kopieren und Verschieben von Dateien in Office365.docx", "removeUrl": null, "websiteUrl": null },
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
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([{
          "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4",
          "displayName": "Document%20Library",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3afbd8cf48-e450-463a-8636-231307dda5f6?label=Document%2520Library&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
          "configuration": {
            "entityId": null,
            "contentUrl": "https://thomy.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente",
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
            "contentUrl": "https://thomy.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente/General/Kopieren und Verschieben von Dateien in Office365.docx",
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

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists all tabs in a Microsoft Teams channel (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs?$expand=teamsApp`) {
        return Promise.resolve({
          value: [{ "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4", "displayName": "Document%20Library", "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3afbd8cf48-e450-463a-8636-231307dda5f6?label=Document%2520Library&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985", "configuration": { "entityId": null, "contentUrl": "https://thomy.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente", "removeUrl": null, "websiteUrl": null }, "teamsApp": { "id": "com.microsoft.teamspace.tab.files.sharepoint", "externalId": null, "displayName": "Document Library", "distributionMethod": "store" } }, { "id": "ba38f554-9ce6-4719-bc9b-e38e4ca16860", "displayName": "CLI-Microsoft365", "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a5a2f05cf-0b6d-4012-b296-342d89158248?webUrl=https%3a%2f%2fgithub.com%2fpnp%2fcli-microsoft365&label=CLI-Microsoft365&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985", "configuration": { "entityId": null, "contentUrl": "https://github.com/pnp/cli-microsoft365", "removeUrl": null, "websiteUrl": "https://github.com/pnp/cli-microsoft365", "dateAdded": "2019-03-28T18:35:53.81Z" }, "teamsApp": { "id": "com.microsoft.teamspace.tab.web", "externalId": null, "displayName": "Website", "distributionMethod": "store" } }, { "id": "b6c511f1-3ad7-4111-8a82-36b13aad4c9e", "displayName": "Word", "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a00b9c26c-b0f4-4c1a-98fa-426aded95ca3?label=Word&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985", "configuration": { "entityId": "C6DBBF49-0290-4194-B3DA-319A72014FD6", "contentUrl": "https://thomy.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente/General/Kopieren und Verschieben von Dateien in Office365.docx", "removeUrl": null, "websiteUrl": null }, "teamsApp": { "id": "com.microsoft.teamspace.tab.file.staticviewer.word", "externalId": null, "displayName": "Word", "distributionMethod": "store" } }, { "id": "d9e972d8-e93d-4b87-beb2-3698912398ea", "displayName": "Wiki", "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3a9d44e015-ae5c-47dc-9577-5af76609e2b0?label=Wiki&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985", "configuration": { "entityId": null, "contentUrl": null, "removeUrl": null, "websiteUrl": null, "wikiTabName": "Wiki", "wikiTabId": 1, "wikiDefaultTab": true }, "teamsApp": { "id": "com.microsoft.teamspace.tab.wiki", "externalId": null, "displayName": "Wiki", "distributionMethod": "store" } }]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000", channelId: "19:00000000000000000000000000000000@thread.skype" } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{
          "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4",
          "displayName": "Document%20Library",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tab%3a%3afbd8cf48-e450-463a-8636-231307dda5f6?label=Document%2520Library&groupId=aa5cf078-46e3-4bcf-9e02-e15ff6efe889&tenantId=a1d5f937-b756-46d7-b92f-464629a6d985",
          "configuration": {
            "entityId": null,
            "contentUrl": "https://thomy.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente",
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
            "contentUrl": "https://thomy.sharepoint.com/sites/MoCaDeSyMo/Freigegebene Dokumente/General/Kopieren und Verschieben von Dateien in Office365.docx",
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});