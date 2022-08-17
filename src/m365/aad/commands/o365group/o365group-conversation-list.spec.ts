import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./o365group-conversation-list');

describe(commands.O365GROUP_CONVERSATION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const jsonOutput = {
    "value": [
      {
        "id": "AAQkAGFhZDhkNGI1LTliZmEtNGEzMi04NTkzLWZjMWExZDkyMWEyZgAQAH4o7SknOTNKqAqMhqJHtUM=",
        "topic": "The new All Company group is ready",
        "hasAttachments": false,
        "lastDeliveredDateTime": "2021-08-02T10:34:00Z",
        "uniqueSenders": [
          "All Company"
        ],
        "preview": "Welcome to the All Company group.Use the group to share ideas, files, and important dates.Start a conversationRead group conversations or start your own.Share filesView, edit, and share all group files, including email attachments.Connect your"
      },
      {
        "id": "AAQkADQzYWUxZTA5LWQwYmItNDcxMy04ZTU5LTg3YmU5NDU3MDlmZgAQAOAzGByAP-dGkFXbXlukzUk=",
        "topic": "Weekly meeting on friday",
        "hasAttachments": false,
        "lastDeliveredDateTime": "2022-02-02T10:34:00Z",
        "uniqueSenders": [
          "John Doe"
        ],
        "preview": "Can we have a weekly meeting on Friday?"
      }
    ]
  };
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
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_CONVERSATION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['topic', 'lastDeliveredDateTime', 'id']);
  });
  it('fails validation if the groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: 'not-c49b-4fd4-8223-28f0ac3a6402' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('Retrieve conversations for the specified group by groupId in the tenant (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/conversations`) {
        return Promise.resolve(
          jsonOutput
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        verbose: true, groupId: "00000000-0000-0000-0000-000000000000"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          jsonOutput.value
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('correctly handles error when listing conversations', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, { options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000" } } as any, (err?: any) => {
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