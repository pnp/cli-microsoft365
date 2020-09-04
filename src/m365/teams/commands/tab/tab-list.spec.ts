import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./tab-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_TAB_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.TEAMS_TAB_LIST), true);
  });

  it('fails validation if the teamId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the channelId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('fails validates for a incorrect channelId missing leading 19:.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validates for a incorrect channelId missing trailing @thread.skpye.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread',
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
      }
    });
    assert.equal(actual, true);
    done();
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
      }
    }, (error?: any) => {
      try {
        assert.equal(JSON.stringify(error), JSON.stringify(new CommandError("Channel id is not in a valid format: 29:d09d9792d59544af846fa19c98b6acc6@thread.skype")));
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          [
            {
              "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4",
              "displayName": "Document%20Library",
              "teamsAppTabId": "com.microsoft.teamspace.tab.files.sharepoint",
            },
            {
              "id": "ba38f554-9ce6-4719-bc9b-e38e4ca16860",
              "displayName": "CLI-Microsoft365",
              "teamsAppTabId": "com.microsoft.teamspace.tab.web",

            },
            {
              "id": "b6c511f1-3ad7-4111-8a82-36b13aad4c9e",
              "displayName": "Word",
              "teamsAppTabId": "com.microsoft.teamspace.tab.file.staticviewer.word",
            },
            {
              "id": "d9e972d8-e93d-4b87-beb2-3698912398ea",
              "displayName": "Wiki",
              "teamsAppTabId": "com.microsoft.teamspace.tab.wiki",
            }
          ]
        ));

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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000", channelId: "19:00000000000000000000000000000000@thread.skype" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "e7cb46d2-b291-409a-b4bc-f5bdd26f10d4",
            "displayName": "Document%20Library",
            "teamsAppTabId": "com.microsoft.teamspace.tab.files.sharepoint",
          },
          {
            "id": "ba38f554-9ce6-4719-bc9b-e38e4ca16860",
            "displayName": "CLI-Microsoft365",
            "teamsAppTabId": "com.microsoft.teamspace.tab.web",

          },
          {
            "id": "b6c511f1-3ad7-4111-8a82-36b13aad4c9e",
            "displayName": "Word",
            "teamsAppTabId": "com.microsoft.teamspace.tab.file.staticviewer.word",
          },
          {
            "id": "d9e972d8-e93d-4b87-beb2-3698912398ea",
            "displayName": "Wiki",
            "teamsAppTabId": "com.microsoft.teamspace.tab.wiki",
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all data in json output mode', (done) => {
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        output: 'json',
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          [
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
        ));

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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TEAMS_TAB_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});