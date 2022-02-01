import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./channel-primary-get');

describe(commands.CHANNEL_PRIMARY_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    Utils.restore([
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
    assert.strictEqual(command.name.startsWith(commands.CHANNEL_PRIMARY_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both teamId and teamName options are not passed', (done) => {
    const actual = command.validate({
      options: {
        channelId: '19:00000000000000000000000000000000@thread.skype'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both teamId and teamName options are passed', (done) => {
    const actual = command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        teamName: 'Team Name'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid guid.', () => {
    const actual = command.validate({
      options: {
        teamId: 'invalid',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    });
    assert.notStrictEqual(actual, true);
  });


  it('fails validation if the teamId is not provided.', (done) => {
    const actual = command.validate({
      options: {
        channelId: '19:00000000000000000000000000000000@thread.skype',
        tabId: '00000000-0000-0000-0000-000000000000'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('correctly validates the when all options are valid', () => {
    const actual = command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails to get team when team does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('The specified team does not exist in the Microsoft Teams');
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

  it('fails when multiple teams with same name exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 2,
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
              "resourceProvisioningOptions": ["Team"]
            },
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
              "resourceProvisioningOptions": ["Team"]
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
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple Microsoft Teams teams with name Team Name found: 00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000`)));
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
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/General?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
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
        assert.strictEqual(call.args[0].displayName, 'General');
        assert.strictEqual(call.args[0].description, null);
        assert.strictEqual(call.args[0].email, '');
        assert.strictEqual(call.args[0].webUrl, 'https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/General?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4');
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

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/primaryChannel`) {
        return Promise.resolve({
          "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
          "displayName": "General",
          "description": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/General?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
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
        assert.strictEqual(call.args[0].displayName, 'General');
        assert.strictEqual(call.args[0].description, null);
        assert.strictEqual(call.args[0].email, '');
        assert.strictEqual(call.args[0].webUrl, 'https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/General?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get primary channel information for the Microsoft Teams team (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/primaryChannel`) {
        return Promise.resolve({
          "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
          "displayName": "General",
          "description": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/General?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
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
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});