import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./tab-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import Sinon = require('sinon');
import * as chalk from 'chalk';

describe(commands.TEAMS_TAB_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
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
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_TAB_ADD), true);
  });

  it('fails validation if the teamId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000',
        appId: 'com.microsoft.teamspace.tab.web',
        appName: 'test',
        contentUrl: '/',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('allows unknown properties', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });

  it('fails validates for a incorrect channelId missing leading 19:.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        appId: 'com.microsoft.teamspace.tab.web',
        appName: 'test',
        contentUrl: '/'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validates for a incorrect channelId missing trailing @thread.skpye.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread',
        appId: 'com.microsoft.teamspace.tab.web',
        appName: 'test',
        contentUrl: '/'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
        appId: 'com.microsoft.teamspace.tab.web',
        appName: 'test',
        contentUrl: '/',
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('creates tab in channel within the Microsoft Teams team in the tenant', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/teams/3b4797e5-bdf3-48e1-a552-839af71562ef`) > -1) {
        return Promise.resolve({
          "id": "19:f3dcbb1674574677abcae89cb626f1e6@thread.skype",
          "displayName": "testweb",
          "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6@thread.skype/"
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamId: '3b4797e5-bdf3-48e1-a552-839af71562ef',
        channelId: '9:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        appId: 'com.microsoft.teamspace.tab.web',
        appName: 'testweb',
        contentUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "id": "19:f3dcbb1674574677abcae89cb626f1e6@thread.skype",
          "displayName": "testweb",
          "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6@thread.skype/"
        }));
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates tab in channel within the Microsoft Teams team in the tenant with all options', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/teams/3b4797e5-bdf3-48e1-a552-839af71562ef`) > -1) {
        return Promise.resolve({
          "id": "19:f3dcbb1674574677abcae89cb626f1e6@thread.skype",
          "displayName": "testweb",
          "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6@thread.skype/"
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamId: '3b4797e5-bdf3-48e1-a552-839af71562ef',
        channelId: '9:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        appId: 'com.microsoft.teamspace.tab.web',
        appName: 'testweb',
        entityId: 'https://xxx.sharepoint.com/Shared%20Documents/',
        removeUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
        contentUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
        websiteUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
        unknown: 'unknown value'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "id": "19:f3dcbb1674574677abcae89cb626f1e6@thread.skype",
          "displayName": "testweb",
          "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6@thread.skype/"
        }));
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('ignores global options when creating request body', (done) => {
    const postStub: Sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`https://graph.microsoft.com/v1.0/teams/3b4797e5-bdf3-48e1-a552-839af71562ef`) > -1) {
        return Promise.resolve({
          "id": "19:f3dcbb1674574677abcae89cb626f1e6@thread.skype",
          "displayName": "testweb",
          "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6@thread.skype/"
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        verbose: true,
        output: "text",
        teamId: '3b4797e5-bdf3-48e1-a552-839af71562ef',
        channelId: '9:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        appId: 'com.microsoft.teamspace.tab.web',
        appName: 'testweb',
        entityId: 'https://xxx.sharepoint.com/Shared%20Documents/',
        removeUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
        contentUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
        websiteUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
        unknown: 'unknown value'
      }
    }, () => {
      try {
        assert.deepEqual(postStub.firstCall.args[0].body, {
          'teamsApp@odata.bind': 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web',
          configuration: {
            contentUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
            entityId: 'https://xxx.sharepoint.com/Shared%20Documents/',
            removeUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
            unknown: 'unknown value',
            websiteUrl: 'https://xxx.sharepoint.com/Shared%20Documents/'
          },
          displayName: 'testweb'
        });
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when adding a tab', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        teamId: '3b4797e5-bdf3-48e1-a552-839af71562ef',
        channelId: '19:eab8fda0837c48edb542574d419ff8ab@thread.skype/tabs',
        appId: 'com.microsoft.teamspace.tab.web',
        appName: 'testweb',
        contentUrl: 'https://xxx.sharepoint.com/Shared%20Documents/',
        websiteUrl: 'https://xxx.sharepoint.com/Shared%20Documents/'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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
});