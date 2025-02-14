import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './m365group-conversation-list.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.M365GROUP_CONVERSATION_LIST, () => {
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
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(true);
    sinon.stub(entraGroup, 'getGroupIdByDisplayName').resolves('00000000-0000-0000-0000-000000000000');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_CONVERSATION_LIST);
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

  it('Retrieve conversations for the group specified by groupId in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/conversations`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true, groupId: "00000000-0000-0000-0000-000000000000"
      }
    });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('Retrieve conversations for the group specified by groupName in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/conversations`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true, groupName: "Finance"
      }
    });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('correctly handles error when listing conversations', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000" } } as any),
      new CommandError('An error has occurred'));
  });

  it('shows error when the group is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { groupId: groupId } } as any),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });

});
