import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './group-list.js';

describe(commands.GROUP_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const groupsFirstBatchList: any = [
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

  const groupsSecondBatchList: any = [
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
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'email', 'privacy', 'external', 'moderated']);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw {
        "error": {
          "base": "An error has occurred."
        }
      };
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('passes validation without parameters', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { limit: 10 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('limit must be a number', async () => {
    const actual = await command.validate({ options: { limit: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('userId must be a number', async () => {
    const actual = await command.validate({ options: { userId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('returns groups without more results', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/groups.json?page=1') {
        return groupsSecondBatchList;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: {} } as any);

    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 4708910);
  });

  it('returns more than 50 groups correctly', async () => {
    let first50Groups: any[] = [];
    // create a batch with 50 groups
    for (let index = 0; index < 25; index++) {
      first50Groups = first50Groups.concat(groupsFirstBatchList);
    }

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/groups.json?page=1') {
        return first50Groups;
      }
      if (opts.url === 'https://www.yammer.com/api/v1/groups.json?page=2') {
        return groupsSecondBatchList;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { output: 'json' } } as any);

    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 53);
  });

  it('returns zero groups when none are found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/groups.json?page=1') {
        return [];
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { output: 'json' } } as any);

    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 0);
  });

  it('returns groups with a specific limit', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      return groupsFirstBatchList;
    });

    await command.action(logger, { options: { limit: 1, output: 'json' } } as any);

    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 1);
  });

  it('handles correct parameters userId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/groups/for_user/10123190123128.json?page=1') {
        return groupsSecondBatchList;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: 10123190123128, output: 'json' } } as any);

    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 4708910);
  });
});
