import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./group-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.YAMMER_GROUP_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let groupsFirstBatchList: any = [
    {
      "type": "group",
      "id": 13114605568,
      "email": "group8+nubo.eu@yammer.com",
      "full_name": "group8",
      "network_id": 801445,
      "name": "group8",
      "description": null,
      "privacy": "public",
      "url": "https://www.yammer.com/api/v1/groups/13114605568",
      "web_url": "https://www.yammer.com/nubo.eu/#/threads/inGroup?type=in_group&feedId=13114605568",
      "mugshot_url": "https://mug0.assets-yammer.com/mugshot/images/48x48/group_profile.png",
      "mugshot_url_template": "https://mug0.assets-yammer.com/mugshot/images/{width}x{height}/group_profile.png",
      "mugshot_id": null,
      "show_in_directory": "true",
      "created_at": "2019/12/21 15:32:09 +0000",
      "color": "#0e4f7a",
      "external": false,
      "moderated": false,
      "header_image_url": "https://mug0.assets-yammer.com/mugshot/images/group-header-megaphone.png",
      "category": "unclassified",
      "default_thread_starter_type": "normal",
      "creator_type": "user",
      "creator_id": 1496550646,
      "state": "active",
      "stats": {
        "members": 1,
        "updates": 0,
        "last_message_id": null,
        "last_message_at": null
      }
    },
    {
      "type": "group",
      "id": 13114597376,
      "email": "group7+nubo.eu@yammer.com",
      "full_name": "group7",
      "network_id": 801445,
      "name": "group7",
      "description": null,
      "privacy": "public",
      "url": "https://www.yammer.com/api/v1/groups/13114597376",
      "web_url": "https://www.yammer.com/nubo.eu/#/threads/inGroup?type=in_group&feedId=13114597376",
      "mugshot_url": "https://mug0.assets-yammer.com/mugshot/images/48x48/group_profile.png",
      "mugshot_url_template": "https://mug0.assets-yammer.com/mugshot/images/{width}x{height}/group_profile.png",
      "mugshot_id": null,
      "show_in_directory": "true",
      "created_at": "2019/12/21 15:32:04 +0000",
      "color": "#0e4f7a",
      "external": false,
      "moderated": false,
      "header_image_url": "https://mug0.assets-yammer.com/mugshot/images/group-header-megaphone.png",
      "category": "unclassified",
      "default_thread_starter_type": "normal",
      "creator_type": "user",
      "creator_id": 1496550646,
      "state": "active",
      "stats": {
        "members": 1,
        "updates": 0,
        "last_message_id": null,
        "last_message_at": null
      }
    }];

  let groupsSecondBatchList: any = [
    {
      "type": "group",
      "id": 4708910,
      "email": "weeklynewscentral+nubo.eu@yammer.com",
      "full_name": "Weekly news Central",
      "network_id": 801445,
      "name": "weeklynewscentral",
      "description": null,
      "privacy": "public",
      "url": "https://www.yammer.com/api/v1/groups/4708910",
      "web_url": "https://www.yammer.com/nubo.eu/#/threads/inGroup?type=in_group&feedId=4708910",
      "mugshot_url": "https://mug0.assets-yammer.com/mugshot/images/48x48/group_profile.png",
      "mugshot_url_template": "https://mug0.assets-yammer.com/mugshot/images/{width}x{height}/group_profile.png",
      "mugshot_id": null,
      "show_in_directory": "true",
      "created_at": "2014/09/12 00:45:34 +0000",
      "color": "#0d5e14",
      "external": false,
      "moderated": false,
      "header_image_url": "https://mug0.assets-yammer.com/mugshot/images/group-header-charts.png",
      "category": "unclassified",
      "default_thread_starter_type": "normal",
      "creator_type": "user",
      "creator_id": 1496550646,
      "state": "active",
      "stats": {
        "members": 2,
        "updates": 0,
        "last_message_id": 510412686,
        "last_message_at": "2015/03/13 11:18:24 +0000"
      }
    },
    {
      "type": "group",
      "id": 4683850,
      "email": "blogboardideas+nubo.eu@yammer.com",
      "full_name": "BLOG Board & Ideas",
      "network_id": 801445,
      "name": "blogboardideas",
      "description": "",
      "privacy": "public",
      "url": "https://www.yammer.com/api/v1/groups/4683850",
      "web_url": "https://www.yammer.com/nubo.eu/#/threads/inGroup?type=in_group&feedId=4683850",
      "mugshot_url": "https://mug0.assets-yammer.com/mugshot/images/48x48/hh6cWz93Qd-sncX8mLhW1sxX1lxZ9mK0",
      "mugshot_url_template": "https://mug0.assets-yammer.com/mugshot/images/{width}x{height}/hh6cWz93Qd-sncX8mLhW1sxX1lxZ9mK0",
      "mugshot_id": "hh6cWz93Qd-sncX8mLhW1sxX1lxZ9mK0",
      "show_in_directory": "true",
      "created_at": "2014/09/02 12:53:39 +0000",
      "color": "#b9555c",
      "external": false,
      "moderated": false,
      "header_image_url": "https://mug0.assets-yammer.com/mugshot/images/group-header-charts.png",
      "category": "unclassified",
      "default_thread_starter_type": "normal",
      "creator_type": "user",
      "creator_id": 1503796713,
      "state": "active",
      "stats": {
        "members": 2,
        "updates": 0,
        "last_message_id": 453774323,
        "last_message_at": "2014/10/20 18:56:01 +0000"
      }
    },
    {
      "type": "group",
      "id": 4742383,
      "email": "nubobenefits+nubo.eu@yammer.com",
      "full_name": "NUBO Benefits",
      "network_id": 801445,
      "name": "nubobenefits",
      "description": null,
      "privacy": "public",
      "url": "https://www.yammer.com/api/v1/groups/4742383",
      "web_url": "https://www.yammer.com/nubo.eu/#/threads/inGroup?type=in_group&feedId=4742383",
      "mugshot_url": "https://mug0.assets-yammer.com/mugshot/images/48x48/group_profile.png",
      "mugshot_url_template": "https://mug0.assets-yammer.com/mugshot/images/{width}x{height}/group_profile.png",
      "mugshot_id": null,
      "show_in_directory": "true",
      "created_at": "2014/09/24 15:01:02 +0000",
      "color": "#446422",
      "external": false,
      "moderated": false,
      "header_image_url": "https://mug0.assets-yammer.com/mugshot/images/group-header-apple.png",
      "category": "unclassified",
      "default_thread_starter_type": "normal",
      "creator_type": "user",
      "creator_id": 1496550646,
      "state": "active",
      "stats": {
        "members": 2,
        "updates": 0,
        "last_message_id": 498516644,
        "last_message_at": "2015/02/13 19:15:43 +0000"
      }
    }];

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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_GROUP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation without parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { limit: 10 } });
    assert.strictEqual(actual, true);
  });

  it('limit must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { limit: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('userId must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { userId: 'abc' } });
    assert.notStrictEqual(actual, true);
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

  it('returns groups without more results', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/groups.json?page=1') {
        return Promise.resolve(groupsSecondBatchList);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 4708910)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns all groups', (done) => {
    let first50Groups: any[] = [];
    // create a batch with 50 groups
    for (let index = 0; index < 25; index++) {
      first50Groups.concat(groupsFirstBatchList);
    }

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/groups.json?page=1') {
        return Promise.resolve(first50Groups);
      }
      if (opts.url === 'https://www.yammer.com/api/v1/groups.json?page=2') {
        return Promise.resolve(groupsSecondBatchList);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].length, 3);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns groups with a specific limit', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve(groupsFirstBatchList);
    });
    cmdInstance.action({ options: { limit: 1, output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].length, 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error in loop', (done) => {
    let i: number = 0;
    let first50Groups: any[] = [];
    // create a batch with 50 groups
    for (let index = 0; index < 25; index++) {
      first50Groups.concat(groupsFirstBatchList);
    }

    sinon.stub(request, 'get').callsFake((opts) => {
      if (i++ === 0) {
        return Promise.resolve(first50Groups);
      }
      else {
        return Promise.reject({
          "error": {
            "base": "An error has occurred."
          }
        });
      }
    });
    cmdInstance.action({ options: { output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles correct parameters userId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/groups/for_user/10123190123128.json?page=1') {
        return Promise.resolve(groupsSecondBatchList);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { userId: 10123190123128, output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 4708910)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});