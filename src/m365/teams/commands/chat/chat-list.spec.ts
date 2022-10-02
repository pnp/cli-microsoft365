import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./chat-list');

describe(commands.CHAT_LIST, () => {
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
    assert.strictEqual(command.name.startsWith(commands.CHAT_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'topic', 'chatType']);
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

  it('fails validation for an incorrect chatType.', async () => {
    const actual = await command.validate({
      options: {
        type: 'oneOn'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input without chat type', async () => {
    const actual = await command.validate({
      options: {
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input for oneOnOne chat conversations', async () => {
    const actual = await command.validate({
      options: {
        type: "oneOnOne"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input for group chat conversations', async () => {
    const actual = await command.validate({
      options: {
        type: "group"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input for meeting chat conversations', async () => {
    const actual = await command.validate({
      options: {
        type: "meeting"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists all chat conversations (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats`) {
        return Promise.resolve({
          "value": [{ "id": "19:meeting_MjdhNjM4YzUtYzExZi00OTFkLTkzZTAtNTVlNmZmMDhkNGU2@thread.v2", "topic": "Meeting chat sample", "createdDateTime": "2020-12-08T23:53:05.801Z", "lastUpdatedDateTime": "2020-12-08T23:58:32.511Z", "chatType": "meeting" }, { "id": "19:561082c0f3f847a58069deb8eb300807@thread.v2", "topic": "Group chat sample", "createdDateTime": "2020-12-03T19:41:07.054Z", "lastUpdatedDateTime": "2020-12-08T23:53:11.012Z", "chatType": "group" }, { "id": "19:d74fc2ed-cb0e-4288-a219-b5c71abaf2aa_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces", "topic": null, "createdDateTime": "2020-12-04T23:10:28.51Z", "lastUpdatedDateTime": "2020-12-04T23:10:36.925Z", "chatType": "oneOnOne" }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        debug: true
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "id": "19:meeting_MjdhNjM4YzUtYzExZi00OTFkLTkzZTAtNTVlNmZmMDhkNGU2@thread.v2",
        "topic": "Meeting chat sample",
        "createdDateTime": "2020-12-08T23:53:05.801Z",
        "lastUpdatedDateTime": "2020-12-08T23:58:32.511Z",
        "chatType": "meeting"
      },
      {
        "id": "19:561082c0f3f847a58069deb8eb300807@thread.v2",
        "topic": "Group chat sample",
        "createdDateTime": "2020-12-03T19:41:07.054Z",
        "lastUpdatedDateTime": "2020-12-08T23:53:11.012Z",
        "chatType": "group"
      },
      {
        "id": "19:d74fc2ed-cb0e-4288-a219-b5c71abaf2aa_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces",
        "topic": null,
        "createdDateTime": "2020-12-04T23:10:28.51Z",
        "lastUpdatedDateTime": "2020-12-04T23:10:36.925Z",
        "chatType": "oneOnOne"
      }
    ]));
  });

  it('lists all chat conversations', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats`) {
        return Promise.resolve({
          "value": [{ "id": "19:meeting_MjdhNjM4YzUtYzExZi00OTFkLTkzZTAtNTVlNmZmMDhkNGU2@thread.v2", "topic": "Meeting chat sample", "createdDateTime": "2020-12-08T23:53:05.801Z", "lastUpdatedDateTime": "2020-12-08T23:58:32.511Z", "chatType": "meeting" }, { "id": "19:561082c0f3f847a58069deb8eb300807@thread.v2", "topic": "Group chat sample", "createdDateTime": "2020-12-03T19:41:07.054Z", "lastUpdatedDateTime": "2020-12-08T23:53:11.012Z", "chatType": "group" }, { "id": "19:d74fc2ed-cb0e-4288-a219-b5c71abaf2aa_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces", "topic": null, "createdDateTime": "2020-12-04T23:10:28.51Z", "lastUpdatedDateTime": "2020-12-04T23:10:36.925Z", "chatType": "oneOnOne" }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        debug: false
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "id": "19:meeting_MjdhNjM4YzUtYzExZi00OTFkLTkzZTAtNTVlNmZmMDhkNGU2@thread.v2",
        "topic": "Meeting chat sample",
        "createdDateTime": "2020-12-08T23:53:05.801Z",
        "lastUpdatedDateTime": "2020-12-08T23:58:32.511Z",
        "chatType": "meeting"
      },
      {
        "id": "19:561082c0f3f847a58069deb8eb300807@thread.v2",
        "topic": "Group chat sample",
        "createdDateTime": "2020-12-03T19:41:07.054Z",
        "lastUpdatedDateTime": "2020-12-08T23:53:11.012Z",
        "chatType": "group"
      },
      {
        "id": "19:d74fc2ed-cb0e-4288-a219-b5c71abaf2aa_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces",
        "topic": null,
        "createdDateTime": "2020-12-04T23:10:28.51Z",
        "lastUpdatedDateTime": "2020-12-04T23:10:36.925Z",
        "chatType": "oneOnOne"
      }
    ]));
  });

  it('lists oneOnOne chat conversations', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats?$filter=chatType eq 'oneOnOne'`) {
        return Promise.resolve({
          "value": [{ "id": "19:d74fc2ed-cb0e-4288-a219-b5c71abaf2aa_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces", "topic": null, "createdDateTime": "2020-12-04T23:10:28.51Z", "lastUpdatedDateTime": "2020-12-04T23:10:36.925Z", "chatType": "oneOnOne" }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        debug: false,
        type: "oneOnOne"
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "id": "19:d74fc2ed-cb0e-4288-a219-b5c71abaf2aa_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces",
        "topic": null,
        "createdDateTime": "2020-12-04T23:10:28.51Z",
        "lastUpdatedDateTime": "2020-12-04T23:10:36.925Z",
        "chatType": "oneOnOne"
      }
    ]));
  });

  it('lists group chat conversations', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats?$filter=chatType eq 'group'`) {
        return Promise.resolve({
          "value": [{ "id": "19:561082c0f3f847a58069deb8eb300807@thread.v2", "topic": "Group chat sample", "createdDateTime": "2020-12-03T19:41:07.054Z", "lastUpdatedDateTime": "2020-12-08T23:53:11.012Z", "chatType": "group" }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        debug: false,
        type: "group"
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "id": "19:561082c0f3f847a58069deb8eb300807@thread.v2",
        "topic": "Group chat sample",
        "createdDateTime": "2020-12-03T19:41:07.054Z",
        "lastUpdatedDateTime": "2020-12-08T23:53:11.012Z",
        "chatType": "group"
      }
    ]));
  });

  it('lists meeting chat conversations', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats?$filter=chatType eq 'meeting'`) {
        return Promise.resolve({
          "value": [{ "id": "19:meeting_MjdhNjM4YzUtYzExZi00OTFkLTkzZTAtNTVlNmZmMDhkNGU2@thread.v2", "topic": "Meeting chat sample", "createdDateTime": "2020-12-08T23:53:05.801Z", "lastUpdatedDateTime": "2020-12-08T23:58:32.511Z", "chatType": "meeting" }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        debug: false,
        type: "meeting"
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "id": "19:meeting_MjdhNjM4YzUtYzExZi00OTFkLTkzZTAtNTVlNmZmMDhkNGU2@thread.v2",
        "topic": "Meeting chat sample",
        "createdDateTime": "2020-12-08T23:53:05.801Z",
        "lastUpdatedDateTime": "2020-12-08T23:58:32.511Z",
        "chatType": "meeting"
      }
    ]));

  });


  it('outputs all data in json output mode', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/chats`) {
        return Promise.resolve({
          "value": [{ "id": "19:meeting_MjdhNjM4YzUtYzExZi00OTFkLTkzZTAtNTVlNmZmMDhkNGU2@thread.v2", "topic": "Meeting chat sample", "createdDateTime": "2020-12-08T23:53:05.801Z", "lastUpdatedDateTime": "2020-12-08T23:58:32.511Z", "chatType": "meeting" }, { "id": "19:561082c0f3f847a58069deb8eb300807@thread.v2", "topic": "Group chat sample", "createdDateTime": "2020-12-03T19:41:07.054Z", "lastUpdatedDateTime": "2020-12-08T23:53:11.012Z", "chatType": "group" }, { "id": "19:d74fc2ed-cb0e-4288-a219-b5c71abaf2aa_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces", "topic": null, "createdDateTime": "2020-12-04T23:10:28.51Z", "lastUpdatedDateTime": "2020-12-04T23:10:36.925Z", "chatType": "oneOnOne" }]
        });
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        debug: false,
        output: 'json'
      }
    });
    assert(loggerLogSpy.calledWith([{ "id": "19:meeting_MjdhNjM4YzUtYzExZi00OTFkLTkzZTAtNTVlNmZmMDhkNGU2@thread.v2", "topic": "Meeting chat sample", "createdDateTime": "2020-12-08T23:53:05.801Z", "lastUpdatedDateTime": "2020-12-08T23:58:32.511Z", "chatType": "meeting" }, { "id": "19:561082c0f3f847a58069deb8eb300807@thread.v2", "topic": "Group chat sample", "createdDateTime": "2020-12-03T19:41:07.054Z", "lastUpdatedDateTime": "2020-12-08T23:53:11.012Z", "chatType": "group" }, { "id": "19:d74fc2ed-cb0e-4288-a219-b5c71abaf2aa_8c0a1a67-50ce-4114-bb6c-da9c5dbcf6ca@unq.gbl.spaces", "topic": null, "createdDateTime": "2020-12-04T23:10:28.51Z", "lastUpdatedDateTime": "2020-12-04T23:10:36.925Z", "chatType": "oneOnOne" }]));
  });

  it('correctly handles error when listing chat conversations', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, { options: { debug: false } } as any), new CommandError('An error has occurred'));
  });
});