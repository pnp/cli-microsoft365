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
import command from './m365group-conversation-post-list.js';
import { settingsNames } from '../../../../settingsNames.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import aadCommands from '../../aadCommands.js';

describe(commands.M365GROUP_CONVERSATION_POST_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const jsonOutput = {
    "value": [
      {
        "id": "AAMkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQBGAAAAAAAItFGwjIkpSKk3RMD2kEsABwB8V4aGbsmzQpcmFTaihptDAAAAAAEMAAB8V4aGbsmzQpcmFTaihptDAAAAABUFAAA=",
        "createdDateTime": "2022-02-21T22:13:53Z",
        "lastModifiedDateTime": "2022-02-21T22:13:53Z",
        "changeKey": "CQAAABYAAAB8V4aGbsmzQpcmFTaihptDAAAAAAKN",
        "categories": [],
        "receivedDateTime": "2022-02-21T22:13:53Z",
        "hasAttachments": false,
        "body": {
          "contentType": "html",
          "content": "<html><body><div>\r\\\n<div dir=\"ltr\">\r\\\n<div dir=\"ltr\">\r\\\n<div style=\"color:black;font-size:12pt;font-family:Calibri,Arial,Helvetica,sans-serif;\">\r\\\nThis is one</div>\r\\\n</div>\r\\\n</div>\r\\\n</div>\r\\\n</body></html>"
        },
        "from": {
          "emailAddress": {
            "name": "Contoso Life",
            "address": "contosolife@M365x435773.onmicrosoft.com"
          }
        },
        "sender": {
          "emailAddress": {
            "name": "Contoso Life",
            "address": "contosolife@M365x435773.onmicrosoft.com"
          }
        }
      },
      {
        "id": "AAMkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQBGAAAAAAAItFGwjIkpSKk3RMD2kEsABwB8V4aGbsmzQpcmFTaihptDAAAAAAEMAAB8V4aGbsmzQpcmFTaihptDAAAAABUGAAA=",
        "createdDateTime": "2022-02-21T22:14:14Z",
        "lastModifiedDateTime": "2022-02-21T22:14:14Z",
        "changeKey": "CQAAABYAAAB8V4aGbsmzQpcmFTaihptDAAAAAAKa",
        "categories": [],
        "receivedDateTime": "2022-02-21T22:14:14Z",
        "hasAttachments": false,
        "body": {
          "contentType": "html",
          "content": "<html><body><div>\r\\\n<div dir=\"ltr\">\r\\\n<div style=\"color:black;font-size:12pt;font-family:Calibri,Arial,Helvetica,sans-serif;\">\r\\\nReply to One</div>\r\\\n</div>\r\\\n</div>\r\\\n</body></html>"
        },
        "from": {
          "emailAddress": {
            "name": "Contoso Life",
            "address": "contosolife@M365x435773.onmicrosoft.com"
          }
        },
        "sender": {
          "emailAddress": {
            "name": "Contoso Life",
            "address": "contosolife@M365x435773.onmicrosoft.com"
          }
        }
      }
    ]
  };
  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(true);
    auth.service.connected = true;
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
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_CONVERSATION_POST_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.M365GROUP_CONVERSATION_POST_LIST]);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['receivedDateTime', 'id']);
  });
  it('fails validation if groupId and groupDisplayName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { groupId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', groupDisplayName: 'MyGroup', threadId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('fails validation if neither groupId nor groupDisplayName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { threadId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('fails validation if the groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: 'not-c49b-4fd4-8223-28f0ac3a6402', threadId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844', threadId: '123' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('Retrieve posts for the specified conversation threadId of m365 group groupId in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/threads/AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E=/posts`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        groupId: "00000000-0000-0000-0000-000000000000",
        threadId: "AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E="
      }
    });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });
  it('Retrieve posts for the specified conversation threadId of m365 group groupDisplayName in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "id": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"
            }
          ]
        };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/threads/AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E=/posts`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        groupDisplayName: "MyGroup",
        threadId: "AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E="
      }
    });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });

  it('correctly handles error when listing posts', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        groupId: "00000000-0000-0000-0000-000000000000",
        threadId: "AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E="
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('shows error when the group is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(entraGroup.isUnifiedGroup);
    sinon.stub(entraGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { groupId: groupId, threadId: 'AAQkADkwN2Q2NDg1LWQ3ZGYtNDViZi1iNGRiLTVhYjJmN2Q5NDkxZQAQAOnRAfDf71lIvrdK85FAn5E=' } } as any),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });
});
