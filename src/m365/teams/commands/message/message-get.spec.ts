import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './message-get.js';

describe(commands.MESSAGE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if teamId, channelId and id are not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if channelId and id are not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171"
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the teamId is not a valid guid', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: "5f5d7b71-1161-44",
      channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
      id: "1540911392778"
    });
    assert.strictEqual(actual.success, false);
  });

  it('validates for a correct input', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
      channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
      id: "1540911392778"
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation for a incorrect channelId missing leading 19:.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '00000000-0000-0000-0000-000000000000',
      channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
      id: "1540911392778"
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation for a incorrect channelId missing trailing @thread.skype.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '00000000-0000-0000-0000-000000000000',
      channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread',
      id: "1540911392778"
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
      channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
      id: "1540911392778",
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('retrieves the specified message (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/5f5d7b71-1161-44d8-bcc1-3da710eb4171/channels/19:88f7e66a8dfe42be92db19505ae912a8@thread.skype/messages/1540911392778`) {
        return {
          attachments: [],
          body: { "contentType": "text", "content": "Konnichiwa" },
          createdDateTime: "2018-10-28T15:56:25.116Z",
          deleted: false,
          etag: "1540742185116",
          from: { "application": null, "device": null, "user": { "id": "c500ecce-645d-4fe1-a2ea-b70f32416b51", "displayName": "Arjen Bloemsma", "identityProvider": "Aad" } },
          id: "1540742185116",
          importance: "normal",
          lastModifiedDateTime: null,
          locale: "en-us",
          mentions: [],
          messageType: "message",
          policyViolation: null,
          reactions: [],
          replyToId: null,
          subject: "",
          summary: null
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
        channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
        id: "1540911392778"
      })
    });
    assert(loggerLogSpy.calledWith({
      attachments: [],
      body: { "contentType": "text", "content": "Konnichiwa" },
      createdDateTime: "2018-10-28T15:56:25.116Z",
      deleted: false,
      etag: "1540742185116",
      from: { "application": null, "device": null, "user": { "id": "c500ecce-645d-4fe1-a2ea-b70f32416b51", "displayName": "Arjen Bloemsma", "identityProvider": "Aad" } },
      id: "1540742185116",
      importance: "normal",
      lastModifiedDateTime: null,
      locale: "en-us",
      mentions: [],
      messageType: "message",
      policyViolation: null,
      reactions: [],
      replyToId: null,
      subject: "",
      summary: null
    }));
  });

  it('retrieves the specified message', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/5f5d7b71-1161-44d8-bcc1-3da710eb4171/channels/19:88f7e66a8dfe42be92db19505ae912a8@thread.skype/messages/1540911392778`) {
        return {
          attachments: [],
          body: { "contentType": "text", "content": "Konnichiwa" },
          createdDateTime: "2018-10-28T15:56:25.116Z",
          deleted: false,
          etag: "1540742185116",
          from: { "application": null, "device": null, "user": { "id": "c500ecce-645d-4fe1-a2ea-b70f32416b51", "displayName": "Arjen Bloemsma", "identityProvider": "Aad" } },
          id: "1540742185116",
          importance: "normal",
          lastModifiedDateTime: null,
          locale: "en-us",
          mentions: [],
          messageType: "message",
          policyViolation: null,
          reactions: [],
          replyToId: null,
          subject: "",
          summary: null
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
        channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
        id: "1540911392778"
      })
    });
    assert(loggerLogSpy.calledWith({
      attachments: [],
      body: { "contentType": "text", "content": "Konnichiwa" },
      createdDateTime: "2018-10-28T15:56:25.116Z",
      deleted: false,
      etag: "1540742185116",
      from: { "application": null, "device": null, "user": { "id": "c500ecce-645d-4fe1-a2ea-b70f32416b51", "displayName": "Arjen Bloemsma", "identityProvider": "Aad" } },
      id: "1540742185116",
      importance: "normal",
      lastModifiedDateTime: null,
      locale: "en-us",
      mentions: [],
      messageType: "message",
      policyViolation: null,
      reactions: [],
      replyToId: null,
      subject: "",
      summary: null
    }));
  });

  it('correctly handles error when retrieving a message', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };

    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
        channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
        id: "1540911392778"
      })
    }), new CommandError('An error has occurred'));
  });
});
