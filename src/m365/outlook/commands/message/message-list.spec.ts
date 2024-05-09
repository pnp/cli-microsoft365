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
import command from './message-list.js';
import { settingsNames } from '../../../../settingsNames.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.MESSAGE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const folderId = 'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=';
  const folderName = 'Inbox';
  const startTime = '2023-12-16';
  const endTime = '2024-01-16';
  const userId = 'fe36f75e-c103-410b-a18a-2bf6df06ac3a';
  const userName = 'john@contoso.com';

  // #region emailOutput
  const emailOutput: any = [
    {
      "id": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALvuv07AAA=",
      "createdDateTime": "2020-09-14T14:30:10Z",
      "lastModifiedDateTime": "2020-09-14T14:30:12Z",
      "changeKey": "CQAAABYAAAAiIsqMbYjsT5e/T7KzowPTAALvssb6",
      "categories": [],
      "receivedDateTime": "2020-09-14T14:30:11Z",
      "sentDateTime": "2020-09-14T14:30:11Z",
      "hasAttachments": false,
      "internetMessageId": "<BY5PR15MB3620A235CD3BCC0C00695BF9CD230@BY5PR15MB3620.namprd15.prod.outlook.com>",
      "subject": "MyAnalytics | Focus Edition",
      "bodyPreview": "MyAnalytics\r\n\r\nDiscover your habits. Work smarter.\r\n\r\nFor your eyes only\r\n\r\nLearn more >\r\n\r\n\r\n\r\nYour month in review: Focus\r\n\r\nDo you have enough uninterrupted time to get your work done?\r\n\r\nAvailable to focus\r\n\r\n73%\r\n\r\n\r\n\r\nCollaboration time\r\n\r\n27%\r\n\r\nNe",
      "importance": "normal",
      "parentFolderId": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=",
      "conversationId": "AAQkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAQAFqY8ThMNbhKoXDeNMiHrJg=",
      "conversationIndex": "AQHWiqONWpjxOEw1uEqhcN40yIesmA==",
      "isDeliveryReceiptRequested": false,
      "isReadReceiptRequested": false,
      "isRead": false,
      "isDraft": false,
      "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e%2FT7KzowPTAAAAAAEMAAAiIsqMbYjsT5e%2FT7KzowPTAALvuv07AAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
      "inferenceClassification": "other",
      "body": {
        "contentType": "html",
        "content": "<html>Hello world</html>"
      },
      "sender": {
        "emailAddress": {
          "name": "MyAnalytics",
          "address": "no-reply@microsoft.com"
        }
      },
      "from": {
        "emailAddress": {
          "name": "MyAnalytics",
          "address": "no-reply@microsoft.com"
        }
      },
      "toRecipients": [
        {
          "emailAddress": {
            "name": "Megan Bowen",
            "address": "MeganB@M365x214355.onmicrosoft.com"
          }
        }
      ],
      "ccRecipients": [],
      "bccRecipients": [],
      "replyTo": [],
      "flag": {
        "flagStatus": "notFlagged"
      }
    },
    {
      "id": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALqongjAAA=",
      "createdDateTime": "2020-09-07T15:45:57Z",
      "lastModifiedDateTime": "2020-09-07T15:45:59Z",
      "changeKey": "CQAAABYAAAAiIsqMbYjsT5e/T7KzowPTAALqm1c9",
      "categories": [],
      "receivedDateTime": "2020-09-07T15:45:57Z",
      "sentDateTime": "2020-09-07T15:45:57Z",
      "hasAttachments": false,
      "internetMessageId": "<BY5PR15MB36206304AB8A6C11B6F99CB6CD280@BY5PR15MB3620.namprd15.prod.outlook.com>",
      "subject": "MyAnalytics | Collaboration Edition",
      "bodyPreview": "MyAnalytics\r\n\r\nDiscover your habits. Work smarter.\r\n\r\nFor your eyes only\r\n\r\nLearn more >\r\n\r\n\r\n\r\nYour month in review: Collaboration\r\n\r\nCould your time working with others be more productive?\r\n\r\n\r\n\r\n27% Collaboration\r\n\r\nin a typical working week\r\n\r\nThis is",
      "importance": "normal",
      "parentFolderId": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=",
      "conversationId": "AAQkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAQAH9vilh-wOZJssjZDqKfAxo=",
      "conversationIndex": "AQHWhS36f2+KWH/A5kmyyNkOop8DGg==",
      "isDeliveryReceiptRequested": false,
      "isReadReceiptRequested": false,
      "isRead": false,
      "isDraft": false,
      "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e%2FT7KzowPTAAAAAAEMAAAiIsqMbYjsT5e%2FT7KzowPTAALqongjAAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
      "inferenceClassification": "other",
      "body": {
        "contentType": "html",
        "content": "<html>Hello world</html>"
      },
      "sender": {
        "emailAddress": {
          "name": "MyAnalytics",
          "address": "no-reply@microsoft.com"
        }
      },
      "from": {
        "emailAddress": {
          "name": "MyAnalytics",
          "address": "no-reply@microsoft.com"
        }
      },
      "toRecipients": [
        {
          "emailAddress": {
            "name": "Megan Bowen",
            "address": "MeganB@M365x214355.onmicrosoft.com"
          }
        }
      ],
      "ccRecipients": [],
      "bccRecipients": [],
      "replyTo": [],
      "flag": {
        "flagStatus": "notFlagged"
      }
    },
    {
      "id": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAALmEVcQAAA=",
      "createdDateTime": "2020-08-31T15:29:21Z",
      "lastModifiedDateTime": "2020-08-31T15:29:23Z",
      "changeKey": "CQAAABYAAAAiIsqMbYjsT5e/T7KzowPTAALmC2Y+",
      "categories": [],
      "receivedDateTime": "2020-08-31T15:29:22Z",
      "sentDateTime": "2020-08-31T15:29:22Z",
      "hasAttachments": false,
      "internetMessageId": "<BY5PR15MB3620FB13B27489F365FECFC1CD510@BY5PR15MB3620.namprd15.prod.outlook.com>",
      "subject": "MyAnalytics | Focus Edition",
      "bodyPreview": "MyAnalytics\r\n\r\nDiscover your habits. Work smarter.\r\n\r\nFor your eyes only\r\n\r\nLearn more >\r\n\r\n\r\n\r\nYour month in review: Focus\r\n\r\nDo you have enough uninterrupted time to get your work done?\r\n\r\nAvailable to focus\r\n\r\n73%\r\n\r\n\r\n\r\nCollaboration time\r\n\r\n27%\r\n\r\nNe",
      "importance": "normal",
      "parentFolderId": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=",
      "conversationId": "AAQkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAQAMCtB8n9pAxGmF3buA8i61g=",
      "conversationIndex": "AQHWf6uAwK0Hyf2kDEaYXdu4DyLrWA==",
      "isDeliveryReceiptRequested": false,
      "isReadReceiptRequested": false,
      "isRead": false,
      "isDraft": false,
      "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e%2FT7KzowPTAAAAAAEMAAAiIsqMbYjsT5e%2FT7KzowPTAALmEVcQAAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
      "inferenceClassification": "other",
      "body": {
        "contentType": "html",
        "content": "<html>Hello world</html>"
      },
      "sender": {
        "emailAddress": {
          "name": "MyAnalytics",
          "address": "no-reply@microsoft.com"
        }
      },
      "from": {
        "emailAddress": {
          "name": "MyAnalytics",
          "address": "no-reply@microsoft.com"
        }
      },
      "toRecipients": [
        {
          "emailAddress": {
            "name": "Megan Bowen",
            "address": "MeganB@M365x214355.onmicrosoft.com"
          }
        }
      ],
      "ccRecipients": [],
      "bccRecipients": [],
      "replyTo": [],
      "flag": {
        "flagStatus": "notFlagged"
      }
    }
  ];
  // #endregion

  const emailResponse: { value: any[] } = { value: emailOutput };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['subject', 'receivedDateTime']);
  });

  it('throws error when using application only permissions and not specifying userId or userName', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('You must specify either the userId or userName option when using app-only permissions.'));
  });

  it('lists messages from the folder with name specified using well-known-name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=100`) {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { folderName: 'inbox' } });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('lists messages from the folder with name specified using well-known-name (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=100`) {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, folderName: 'inbox' } });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('lists messages from the folder with id specified using well-known-name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=100`) {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { folderId: 'inbox' } });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('lists messages from the folder with the specified name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders/AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=/messages?$top=100`) {
        return emailResponse;
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders?$filter=displayName eq 'SecondInbox'&$select=id`) {
        return {
          "value": [
            {
              "id": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA="
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { folderName: 'SecondInbox' } });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('lists messages from the folder with the specified id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders/AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=/messages?$top=100`) {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { folderId: 'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=' } });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('lists messages from the currently logged in user', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/messages?$top=100`) {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('lists messages for the currently logged in user with a specified startTime', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/messages?$top=100&$filter=receivedDateTime ge ${startTime}`) {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { startTime: startTime } });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('lists messages for the currently logged in user with a specified endTime and specifying a user by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/messages?$top=100&$filter=receivedDateTime lt ${endTime}`) {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId, endTime: endTime } });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('lists messages for the currently logged in user with a specified start and endTime and specifying a user by UPN', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName)}/messages?$top=100&$filter=receivedDateTime ge ${startTime} and receivedDateTime lt ${endTime}`) {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { startTime: startTime, endTime: endTime, userName: userName } });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('returns error when the folder with the specified name does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders?$filter=displayName eq 'Imbox'&$select=id`) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { folderName: 'Imbox' } } as any),
      new CommandError(`Folder with name 'Imbox' not found`));
  });

  it('returns error when multiple folders with the specified name found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders?$filter=displayName eq 'Archives'&$select=id`) {
        return {
          "value": [
            {
              "id": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA="
            },
            {
              "id": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAB="
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { folderName: 'Archives' } } as any),
      new CommandError("Multiple folders with name 'Archives' found. Found: AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=, AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAB=."));
  });

  it('handles selecting single result when multiple folders with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders?$filter=displayName eq 'Archives'&$select=id`) {
        return {
          "value": [
            {
              "id": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA="
            },
            {
              "id": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAB="
            }
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailFolders/AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA=/messages?$top=100') {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({
      "id": "AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAA="
    });

    await command.action(logger, { options: { folderName: 'Archives' } });
    assert(loggerLogSpy.calledOnceWith(emailOutput));
  });

  it('returns all message properties in JSON output mode', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=100`) {
        return emailResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { folderName: 'inbox', output: 'json' } });
    assert(loggerLogSpy.calledWith(emailResponse.value));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('An error has occurred'));
  });

  it('passes validation if both start and endTime are valid ISO datetimes', async () => {
    const actual = await command.validate({ options: { startTime: startTime, endTime: endTime } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if startTime is not a valid ISO datetime', async () => {
    const actual = await command.validate({ options: { startTime: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if endTime is not a valid ISO datetime', async () => {
    const actual = await command.validate({ options: { endTime: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if endTime is in the future', async () => {
    const endTime = new Date();
    endTime.setHours(endTime.getHours() + 1);
    const actual = await command.validate({ options: { endTime: endTime.toISOString() } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if startTime is in the future', async () => {
    const startTime = new Date();
    startTime.setHours(startTime.getHours() + 1);
    const actual = await command.validate({ options: { startTime: startTime.toISOString() } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if endTime is before startTime', async () => {
    const startTime = new Date();
    const endTime = new Date(startTime);
    endTime.setTime(endTime.getTime() - 1);
    const actual = await command.validate({ options: { startTime: startTime.toISOString(), endTime: endTime.toISOString() } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both folderId and folderName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { folderId: folderId, folderName: folderName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid user principal name', async () => {
    const actual = await command.validate({ options: { userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if userName is a valid user principal name', async () => {
    const actual = await command.validate({ options: { userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
