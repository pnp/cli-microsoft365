import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./channel-get');

describe(commands.CHANNEL_GET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.CHANNEL_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both teamId and teamName options are not passed', async () => {
    const actual = await command.validate({
      options: {
        channelId: '19:00000000000000000000000000000000@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both teamId and teamName options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        teamName: 'Team Name',
        channelId: '19:00000000000000000000000000000000@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelId, channelName and primary options are not passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelId with primary options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        primary: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelName and primary options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelName: 'Channel Name',
        primary: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelId, channelName and primary options are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        channelName: 'Channel Name',
        primary: true
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelId and channelName are passed', async () => {
    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        channelName: 'Channel Name'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not provided.', async () => {
    const actual = await command.validate({
      options: {
        channelId: '19:00000000000000000000000000000000@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the channelId is not provided.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validates for a incorrect channelId missing leading 19:.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00000000000000000000000000000000@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validates for a incorrect channelId missing trailing @thread.skype.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly validates the when all options are valid', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails to get channel information due to wrong channel id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/channels/19%3A493665404ebd4a18adb8a980a31b4986%40thread.skype`) {
        return Promise.reject({
          "error": {
            "code": "ItemNotFound",
            "message": "Failed to execute Skype backend request GetThreadS2SRequest.",
            "innerError": {
              "request-id": "4bebd0d2-d154-491b-b73f-d59ad39646fb",
              "date": "2019-04-06T13:40:51"
            }
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamId: '39958f28-eefb-4006-8f83-13b6ac2a4a7f',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Failed to execute Skype backend request GetThreadS2SRequest.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when team name does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
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
              "discoverySettings": null,
              "resourceProvisioningOptions": []
            }
          ]
        }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name',
        channelName: 'Channel Name',
        tabName: 'Tab Name'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified team does not exist in the Microsoft Teams`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get channel when channel does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/channels?$filter=displayName eq '`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: 'Channel Name',
        tabName: 'Tab Name'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified channel does not exist in the Microsoft Teams team`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get channel information for the Microsoft Teams team by id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/channels/19%3A493665404ebd4a18adb8a980a31b4986%40thread.skype`) {
        return Promise.resolve({
          "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
          "displayName": "channel1",
          "description": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        teamId: '39958f28-eefb-4006-8f83-13b6ac2a4a7f',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
        assert.strictEqual(call.args[0].id, '19:493665404ebd4a18adb8a980a31b4986@thread.skype');
        assert.strictEqual(call.args[0].displayName, 'channel1');
        assert.strictEqual(call.args[0].description, null);
        assert.strictEqual(call.args[0].email, '');
        assert.strictEqual(call.args[0].webUrl, 'https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get primary channel information for the Microsoft Teams team by id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/primaryChannel`) {
        return Promise.resolve({
          "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
          "displayName": "General",
          "description": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/general?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        teamId: '39958f28-eefb-4006-8f83-13b6ac2a4a7f',
        primary: true
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
        assert.strictEqual(call.args[0].id, '19:493665404ebd4a18adb8a980a31b4986@thread.skype');
        assert.strictEqual(call.args[0].displayName, 'General');
        assert.strictEqual(call.args[0].description, null);
        assert.strictEqual(call.args[0].email, '');
        assert.strictEqual(call.args[0].webUrl, 'https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/general?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get channel information for the Microsoft Teams team by name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "39958f28-eefb-4006-8f83-13b6ac2a4a7f",
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
              "discoverySettings": null,
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }

      if ((opts.url as string).indexOf(`/channels?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
              "displayName": "Channel Name",
              "description": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4",
              "membershipType": "standard"
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/channels/19%3A493665404ebd4a18adb8a980a31b4986%40thread.skype`) {
        return Promise.resolve({
          "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
          "displayName": "Channel Name",
          "description": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        teamName: 'Team Name',
        channelName: 'Channel Name'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
        assert.strictEqual(call.args[0].id, '19:493665404ebd4a18adb8a980a31b4986@thread.skype');
        assert.strictEqual(call.args[0].displayName, 'Channel Name');
        assert.strictEqual(call.args[0].description, null);
        assert.strictEqual(call.args[0].email, '');
        assert.strictEqual(call.args[0].webUrl, 'https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get primary channel information for the Microsoft Teams team by name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "39958f28-eefb-4006-8f83-13b6ac2a4a7f",
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
              "discoverySettings": null,
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }

      if ((opts.url as string).indexOf(`/channels?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
              "displayName": "General",
              "description": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/general?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4",
              "membershipType": "standard"
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/primaryChannel`) {
        return Promise.resolve({
          "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
          "displayName": "General",
          "description": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/general?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        teamName: 'Team Name',
        primary: true
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = loggerLogSpy.lastCall;
        assert.strictEqual(call.args[0].id, '19:493665404ebd4a18adb8a980a31b4986@thread.skype');
        assert.strictEqual(call.args[0].displayName, 'General');
        assert.strictEqual(call.args[0].description, null);
        assert.strictEqual(call.args[0].email, '');
        assert.strictEqual(call.args[0].webUrl, 'https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/general?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get channel information for the Microsoft Teams team (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/channels/19%3A493665404ebd4a18adb8a980a31b4986%40thread.skype`) {
        return Promise.resolve({
          "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
          "displayName": "channel1",
          "description": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamId: '39958f28-eefb-4006-8f83-13b6ac2a4a7f',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = loggerLogSpy.getCall(loggerLogSpy.callCount - 2);
        assert.strictEqual(call.args[0].id, '19:493665404ebd4a18adb8a980a31b4986@thread.skype');
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