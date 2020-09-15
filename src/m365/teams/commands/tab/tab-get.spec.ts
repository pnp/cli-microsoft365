import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./tab-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_TAB_GET, () => {
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
    assert.equal(command.name.startsWith(commands.TEAMS_TAB_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('fails validation if both teamId and teamName options are not passed', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if both teamId and teamName options are passed', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        teamName: 'Team Name',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if both channelId and channelName options are not passed', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if both channelId and channelName options are passed', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        channelName: 'Channel Name',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if both tabId and tabName options are not passed', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if both tabId and tabName options are passed', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07',
        tabName: 'Tab Name'
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
        tabId: '00000000-0000-0000-0000-000000000000'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '00000000-0000-0000-0000-000000000000'
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


  it('fails validation if the tabId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '00000000-0000'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the tabId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype'
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
        tabName: 'Tab Name'
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
      done();
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '29:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        tabId: '1432c9da-8b9c-4602-9248-e0800f3e3f07'
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

  it('should get team id of a Microsoft Teams Tab by team name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams?$filter=displayName eq 'Team Name'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
            {
              "id": "a3e044e8-7532-44a2-90d4-fe4ac19bc9a6",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null
            }
          ]
        });
      }
      done();
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamName: 'Team Name',
        channelName: 'Channel Name',
        tabName: 'Tab Name'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          "a3e044e8-7532-44a2-90d4-fe4ac19bc9a6"
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get channel id of a Microsoft Teams Tab by channel name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels?$filter=displayName eq 'Channel Name'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('340537cc-f8e0-41f8-97b0-6cf0218d3357')/channels",
          "@odata.count": 1,
          "value": [
            {
              "id": "19:a169ff3b9bb344e382e0fb7972826e1c@thread.tacv2",
              "displayName": "Channel Name",
              "description": "Channel description",
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3aa169ff3b9bb344e382e0fb7972826e1c%40thread.tacv2/General?groupId=340537cc-f8e0-41f8-97b0-6cf0218d3357&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
              "membershipType": "standard"
            }
          ]
        });
      }
      done();
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamName: 'Team Name',
        channelName: 'Channel Name',
        tabName: 'Tab Name'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          "19:a169ff3b9bb344e382e0fb7972826e1c@thread.tacv2"
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get tab id of a Microsoft Teams Tab by tab name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs?$filter=displayName eq 'Tab Name'`) {
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
          }
        });
      }
      done();
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamName: 'Team Name',
        channelName: 'Channel Name',
        tabName: 'Tab Name'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          "00000000-0000-0000-0000-000000000000"
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get url of a Microsoft Teams Tab', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs/00000000-0000-0000-0000-000000000000`) {
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
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        tabId: '00000000-0000-0000-0000-000000000000'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          "https://teams.microsoft.com/l/entity/4d3b7fcd-b601-4718-9021-b88dbab77e26/_djb2_msteams_prefix_b1d6cbec-fb9d-4d5f-996c-b65abcd13bef?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3a7b6aabe5c04d4a12b813f9272b0774f8%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=a3e044e8-7532-44a2-90d4-fe4ac19bc9a6&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4"
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
    assert(find.calledWith(commands.TEAMS_TAB_GET));
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