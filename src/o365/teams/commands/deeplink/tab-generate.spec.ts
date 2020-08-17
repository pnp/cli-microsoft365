import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./tab-generate');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_DEEPLINK_TAB_GENERATE, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    assert.equal(command.name.startsWith(commands.TEAMS_DEEPLINK_TAB_GENERATE), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('fails validation if the teamId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        label: 'work',
        tabType: 'Configurable'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        label: 'work',
        tabType: 'Configurable'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the channelId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        label: 'work',
        tabType: 'Configurable'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the channelId is not valid channelId', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: 'invalid',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        label: 'work',
        tabType: 'Configurable'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validates for a incorrect channelId missing leading 19:.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        label: 'work',
        tabType: 'Configurable'
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
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        label: 'work',
        tabType: 'Configurable'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the tabId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        label: 'work',
        tabType: 'Configurable'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the tabId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '00000000-0000',
        label: 'work',
        tabType: 'Configurable'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the label is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        tabType: 'Configurable'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the tabType is not valid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        label: 'work',
        tabType: 'Invalid'
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
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        label: 'work',
        tabType: 'Configurable'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('correctly handles teams tabs request failure due to wrong channel id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/29%3A552b7125655c46d5b5b86db02ee7bfdf%40thread.skype/tabs/1432c9da-8b9c-4602-9248-e0800f3e3f07?$expand=teamsApp`) {
        return Promise.reject({
          "error": {
            "code": "Invalid request",
            "message": "Channel id is not in a valid format: 29:552b7125655c46d5b5b86db02ee7bfdf@thread.skype",
            "innerError": {
              "request-id": "75c4e0f1-035e-47e3-917b-0c8823a02a96",
              "date": "2020-07-19T11:08:32"
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
        channelId: '29:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        label: 'work',
        tabType: 'Configurable'
      }
    }, (error?: any) => {
      try {
        assert.equal(JSON.stringify(error), JSON.stringify(new CommandError("Channel id is not in a valid format: 29:552b7125655c46d5b5b86db02ee7bfdf@thread.skype")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get deeplink for tab in a Microsoft Teams channel with Configurable tabType', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs/00000000-0000-0000-0000-000000000000?$expand=teamsApp`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('a3e044e8-7532-44a2-90d4-fe4ac19bc9a6')/channels('19%3A7b6aabe5c04d4a12b813f9272b0774f8%40thread.skype')/tabs(teamsApp())/$entity",
          "id": "1432c9da-8b9c-4602-9248-e0800f3e3f07",
          "displayName": "TeamsTab",
          "webUrl": "https://teams.microsoft.com/l/entity/4d3b7fcd-b601-4718-9021-b88dbab77e26/_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3a7b6aabe5c04d4a12b813f9272b0774f8%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=a3e044e8-7532-44a2-90d4-fe4ac19bc9a6&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
          "configuration": {
            "entityId": "sharepointtab_ddfbc744-622f-4214-98a0-e276ef32d351",
            "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
            "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
            "websiteUrl": null,
            "dateAdded": "2020-07-18T19:27:22.03Z"
          },
          "teamsApp": {
            "id": "4d3b7fcd-b601-4718-9021-b88dbab77e26",
            "externalId": "0172ff63-158d-44b5-aa23-99e72a812c02",
            "displayName": "TeamsTab",
            "distributionMethod": "organization"
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
        tabId: '00000000-0000-0000-0000-000000000000',
        label: 'work',
        tabType: 'Configurable'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = cmdInstanceLogSpy.lastCall;
        assert.equal(call.args[0].deeplink, `https://teams.microsoft.com/l/entity/4d3b7fcd-b601-4718-9021-b88dbab77e26/_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef?webUrl=https%3A%2F%2Fteams.microsoft.com%2Fl%2Fentity%2F4d3b7fcd-b601-4718-9021-b88dbab77e26%2F_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef%3Flabel%3DTeamsTab%26context%3D%257b%250d%250a%2B%2B%2522canvasUrl%2522%253a%2B%2522https%253a%252f%252fcontoso.sharepoint.com%252fsites%252fPrototypeTeam%252f_layouts%252f15%252fTeamsLogon.aspx%253fSPFX%253dtrue%2526dest%253d%252fsites%252fPrototypeTeam%252f_layouts%252f15%252fteamshostedapp.aspx%25253Flist%253d7d7f911a-bf19-46a0-86d9-187c3f32cce2%252526id%253d2%252526webPartInstanceId%253d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%2522%252c%250d%250a%2B%2B%2522channelId%2522%253a%2B%252219%253a7b6aabe5c04d4a12b813f9272b0774f8%2540thread.skype%2522%252c%250d%250a%2B%2B%2522subEntityId%2522%253a%2Bnull%250d%250a%257d%26groupId%3Da3e044e8-7532-44a2-90d4-fe4ac19bc9a6%26tenantId%3Dde348bc7-1aeb-4406-8cb3-97db021cadb4&label=work&context={"channelId": "19%3A00000000000000000000000000000000%40thread.skype"}`);
        
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get deeplink for tab in a Microsoft Teams channel with Static tabType', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs/00000000-0000-0000-0000-000000000000?$expand=teamsApp`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('a3e044e8-7532-44a2-90d4-fe4ac19bc9a6')/channels('19%3A7b6aabe5c04d4a12b813f9272b0774f8%40thread.skype')/tabs(teamsApp())/$entity",
          "id": "1432c9da-8b9c-4602-9248-e0800f3e3f07",
          "displayName": "TeamsTab",
          "webUrl": "https://teams.microsoft.com/l/entity/4d3b7fcd-b601-4718-9021-b88dbab77e26/_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3a7b6aabe5c04d4a12b813f9272b0774f8%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=a3e044e8-7532-44a2-90d4-fe4ac19bc9a6&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
          "configuration": {
            "entityId": "sharepointtab_ddfbc744-622f-4214-98a0-e276ef32d351",
            "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
            "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
            "websiteUrl": null,
            "dateAdded": "2020-07-18T19:27:22.03Z"
          },
          "teamsApp": {
            "id": "4d3b7fcd-b601-4718-9021-b88dbab77e26",
            "externalId": "0172ff63-158d-44b5-aa23-99e72a812c02",
            "displayName": "TeamsTab",
            "distributionMethod": "organization"
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
        tabId: '00000000-0000-0000-0000-000000000000',
        label: 'work',
        tabType: 'Static'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = cmdInstanceLogSpy.lastCall;
        assert.equal(call.args[0].deeplink, "https://teams.microsoft.com/l/entity/4d3b7fcd-b601-4718-9021-b88dbab77e26/_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef?webUrl=https%3A%2F%2Fteams.microsoft.com%2Fl%2Fentity%2F4d3b7fcd-b601-4718-9021-b88dbab77e26%2F_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef%3Flabel%3DTeamsTab%26context%3D%257b%250d%250a%2B%2B%2522canvasUrl%2522%253a%2B%2522https%253a%252f%252fcontoso.sharepoint.com%252fsites%252fPrototypeTeam%252f_layouts%252f15%252fTeamsLogon.aspx%253fSPFX%253dtrue%2526dest%253d%252fsites%252fPrototypeTeam%252f_layouts%252f15%252fteamshostedapp.aspx%25253Flist%253d7d7f911a-bf19-46a0-86d9-187c3f32cce2%252526id%253d2%252526webPartInstanceId%253d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%2522%252c%250d%250a%2B%2B%2522channelId%2522%253a%2B%252219%253a7b6aabe5c04d4a12b813f9272b0774f8%2540thread.skype%2522%252c%250d%250a%2B%2B%2522subEntityId%2522%253a%2Bnull%250d%250a%257d%26groupId%3Da3e044e8-7532-44a2-90d4-fe4ac19bc9a6%26tenantId%3Dde348bc7-1aeb-4406-8cb3-97db021cadb4&label=work");
        
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get deeplink for tab in a Microsoft Teams channel with default tabType', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs/00000000-0000-0000-0000-000000000000?$expand=teamsApp`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('a3e044e8-7532-44a2-90d4-fe4ac19bc9a6')/channels('19%3A7b6aabe5c04d4a12b813f9272b0774f8%40thread.skype')/tabs(teamsApp())/$entity",
          "id": "1432c9da-8b9c-4602-9248-e0800f3e3f07",
          "displayName": "TeamsTab",
          "webUrl": "https://teams.microsoft.com/l/entity/4d3b7fcd-b601-4718-9021-b88dbab77e26/_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3a7b6aabe5c04d4a12b813f9272b0774f8%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=a3e044e8-7532-44a2-90d4-fe4ac19bc9a6&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
          "configuration": {
            "entityId": "sharepointtab_ddfbc744-622f-4214-98a0-e276ef32d351",
            "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
            "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
            "websiteUrl": null,
            "dateAdded": "2020-07-18T19:27:22.03Z"
          },
          "teamsApp": {
            "id": "4d3b7fcd-b601-4718-9021-b88dbab77e26",
            "externalId": "0172ff63-158d-44b5-aa23-99e72a812c02",
            "displayName": "TeamsTab",
            "distributionMethod": "organization"
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
        tabId: '00000000-0000-0000-0000-000000000000',
        label: 'work',
        tabType: 'Static'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = cmdInstanceLogSpy.lastCall;
        assert.equal(call.args[0].deeplink, "https://teams.microsoft.com/l/entity/4d3b7fcd-b601-4718-9021-b88dbab77e26/_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef?webUrl=https%3A%2F%2Fteams.microsoft.com%2Fl%2Fentity%2F4d3b7fcd-b601-4718-9021-b88dbab77e26%2F_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef%3Flabel%3DTeamsTab%26context%3D%257b%250d%250a%2B%2B%2522canvasUrl%2522%253a%2B%2522https%253a%252f%252fcontoso.sharepoint.com%252fsites%252fPrototypeTeam%252f_layouts%252f15%252fTeamsLogon.aspx%253fSPFX%253dtrue%2526dest%253d%252fsites%252fPrototypeTeam%252f_layouts%252f15%252fteamshostedapp.aspx%25253Flist%253d7d7f911a-bf19-46a0-86d9-187c3f32cce2%252526id%253d2%252526webPartInstanceId%253d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%2522%252c%250d%250a%2B%2B%2522channelId%2522%253a%2B%252219%253a7b6aabe5c04d4a12b813f9272b0774f8%2540thread.skype%2522%252c%250d%250a%2B%2B%2522subEntityId%2522%253a%2Bnull%250d%250a%257d%26groupId%3Da3e044e8-7532-44a2-90d4-fe4ac19bc9a6%26tenantId%3Dde348bc7-1aeb-4406-8cb3-97db021cadb4&label=work");
        
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('should get deeplink for tab in a Microsoft Teams channel with default tabType (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs/00000000-0000-0000-0000-000000000000?$expand=teamsApp`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('a3e044e8-7532-44a2-90d4-fe4ac19bc9a6')/channels('19%3A7b6aabe5c04d4a12b813f9272b0774f8%40thread.skype')/tabs(teamsApp())/$entity",
          "id": "1432c9da-8b9c-4602-9248-e0800f3e3f07",
          "displayName": "TeamsTab",
          "webUrl": "https://teams.microsoft.com/l/entity/4d3b7fcd-b601-4718-9021-b88dbab77e26/_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3a7b6aabe5c04d4a12b813f9272b0774f8%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=a3e044e8-7532-44a2-90d4-fe4ac19bc9a6&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
          "configuration": {
            "entityId": "sharepointtab_ddfbc744-622f-4214-98a0-e276ef32d351",
            "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
            "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
            "websiteUrl": null,
            "dateAdded": "2020-07-18T19:27:22.03Z"
          },
          "teamsApp": {
            "id": "4d3b7fcd-b601-4718-9021-b88dbab77e26",
            "externalId": "0172ff63-158d-44b5-aa23-99e72a812c02",
            "displayName": "TeamsTab",
            "distributionMethod": "organization"
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        output: 'json',
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        tabId: '00000000-0000-0000-0000-000000000000',
        label: 'work'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "deeplink": "https://teams.microsoft.com/l/entity/4d3b7fcd-b601-4718-9021-b88dbab77e26/_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef?webUrl=https%3A%2F%2Fteams.microsoft.com%2Fl%2Fentity%2F4d3b7fcd-b601-4718-9021-b88dbab77e26%2F_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef%3Flabel%3DTeamsTab%26context%3D%257b%250d%250a%2B%2B%2522canvasUrl%2522%253a%2B%2522https%253a%252f%252fcontoso.sharepoint.com%252fsites%252fPrototypeTeam%252f_layouts%252f15%252fTeamsLogon.aspx%253fSPFX%253dtrue%2526dest%253d%252fsites%252fPrototypeTeam%252f_layouts%252f15%252fteamshostedapp.aspx%25253Flist%253d7d7f911a-bf19-46a0-86d9-187c3f32cce2%252526id%253d2%252526webPartInstanceId%253d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%2522%252c%250d%250a%2B%2B%2522channelId%2522%253a%2B%252219%253a7b6aabe5c04d4a12b813f9272b0774f8%2540thread.skype%2522%252c%250d%250a%2B%2B%2522subEntityId%2522%253a%2Bnull%250d%250a%257d%26groupId%3Da3e044e8-7532-44a2-90d4-fe4ac19bc9a6%26tenantId%3Dde348bc7-1aeb-4406-8cb3-97db021cadb4&label=work"
        }));
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
    assert(find.calledWith(commands.TEAMS_DEEPLINK_TAB_GENERATE));
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