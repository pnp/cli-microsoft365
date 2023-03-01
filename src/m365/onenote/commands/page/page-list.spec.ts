import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { odata } from '../../../../utils/odata';
import { formatting } from '../../../../utils/formatting';
const command: Command = require('./page-list');

describe(commands.PAGE_LIST, () => {
  const userId = '0e38b3b3-d9ac-42fa-81db-437ac8caec2f';
  const userName = 'john@contoso.com';
  const groupId = 'bba4c915-0ac8-47a1-bd05-087a44c92d3b';
  const groupName = 'Dummy Group A';
  const webUrl = 'https://contoso.sharepoint.com/sites/HR';
  const siteId = 'contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2';
  const pageResponse = {
    value: [
      {
        "id": "1-a26aaec43ed348bd82edf4eb44e73d6c!14-3eb21088-b613-4698-98df-92a7d34e0678",
        "self": "https://graph.microsoft.com/v1.0/users/mathijs@mathijsdev2.onmicrosoft.com/onenote/pages/1-a26aaec43ed348bd82edf4eb44e73d6c!14-3eb21088-b613-4698-98df-92a7d34e0678",
        "createdDateTime": "2023-01-07T10:56:45Z",
        "title": "Page A",
        "createdByAppId": "",
        "contentUrl": "https://graph.microsoft.com/v1.0/users/mathijs@mathijsdev2.onmicrosoft.com/onenote/pages/1-a26aaec43ed348bd82edf4eb44e73d6c!14-3eb21088-b613-4698-98df-92a7d34e0678/content",
        "lastModifiedDateTime": "2023-01-07T10:57:24Z",
        "links": {
          "oneNoteClientUrl": {
            "href": "onenote:https://mathijsdev2-my.sharepoint.com/personal/mathijs_mathijsdev2_onmicrosoft_com/Documents/Notitieblokken/My%20OneNote/Test.one#Page%20A&section-id=94cacaca-d6b5-428d-b967-d3cf01b95c28&page-id=8ca085ba-cad9-4cf4-824e-07e66520ac3f&end"
          },
          "oneNoteWebUrl": {
            "href": "https://mathijsdev2-my.sharepoint.com/personal/mathijs_mathijsdev2_onmicrosoft_com/Documents/Notitieblokken/My%20OneNote?wd=target%28Test.one%7C94cacaca-d6b5-428d-b967-d3cf01b95c28%2FPage%20A%7C8ca085ba-cad9-4cf4-824e-07e66520ac3f%2F%29"
          }
        },
        "parentSection": {
          "id": "1-3eb21088-b613-4698-98df-92a7d34e0678",
          "displayName": "Test",
          "self": "https://graph.microsoft.com/v1.0/users/mathijs@mathijsdev2.onmicrosoft.com/onenote/sections/1-3eb21088-b613-4698-98df-92a7d34e0678"
        }
      },
      {
        "id": "1-a26aaec43ed348bd82edf4eb44e73d6c!68-3eb21088-b613-4698-98df-92a7d34e0678",
        "self": "https://graph.microsoft.com/v1.0/users/mathijs@mathijsdev2.onmicrosoft.com/onenote/pages/1-a26aaec43ed348bd82edf4eb44e73d6c!68-3eb21088-b613-4698-98df-92a7d34e0678",
        "createdDateTime": "2023-01-07T10:57:15Z",
        "title": "Page B",
        "createdByAppId": "",
        "contentUrl": "https://graph.microsoft.com/v1.0/users/mathijs@mathijsdev2.onmicrosoft.com/onenote/pages/1-a26aaec43ed348bd82edf4eb44e73d6c!68-3eb21088-b613-4698-98df-92a7d34e0678/content",
        "lastModifiedDateTime": "2023-01-07T10:57:17Z",
        "links": {
          "oneNoteClientUrl": {
            "href": "onenote:https://mathijsdev2-my.sharepoint.com/personal/mathijs_mathijsdev2_onmicrosoft_com/Documents/Notitieblokken/My%20OneNote/Test.one#Page%20B&section-id=94cacaca-d6b5-428d-b967-d3cf01b95c28&page-id=46a1b220-7ffd-4512-a571-55322097c08d&end"
          },
          "oneNoteWebUrl": {
            "href": "https://mathijsdev2-my.sharepoint.com/personal/mathijs_mathijsdev2_onmicrosoft_com/Documents/Notitieblokken/My%20OneNote?wd=target%28Test.one%7C94cacaca-d6b5-428d-b967-d3cf01b95c28%2FPage%20B%7C46a1b220-7ffd-4512-a571-55322097c08d%2F%29"
          }
        },
        "parentSection": {
          "id": "1-3eb21088-b613-4698-98df-92a7d34e0678",
          "displayName": "Test",
          "self": "https://graph.microsoft.com/v1.0/users/mathijs@mathijsdev2.onmicrosoft.com/onenote/sections/1-3eb21088-b613-4698-98df-92a7d34e0678"
        }
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      request.get,
      odata.getAllItems
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['createdDateTime', 'title', 'id']);
  });

  it('fails validation if the userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: groupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if no option specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists Microsoft OneNote pages for the currently logged in user', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/me/onenote/pages`) {
        return pageResponse.value;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(pageResponse.value));
  });

  it('lists Microsoft OneNote pages for user by id', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/users/${userId}/onenote/pages`) {
        return pageResponse.value;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { userId: userId } });
    assert(loggerLogSpy.calledWith(pageResponse.value));
  });

  it('lists Microsoft OneNote pages for user by name', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/users/${userName}/onenote/pages`) {
        return pageResponse.value;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { userName: userName } });
    assert(loggerLogSpy.calledWith(pageResponse.value));
  });

  it('lists Microsoft OneNote pages in group by id', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/groups/${groupId}/onenote/pages`) {
        return pageResponse.value;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupId: groupId } });
    assert(loggerLogSpy.calledWith(pageResponse.value));
  });

  it('lists Microsoft OneNote pages in group by name', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(groupName)}'`) {
        return [{
          "id": groupId,
          "description": groupName,
          "displayName": groupName
        }];
      }
      if (url === `https://graph.microsoft.com/v1.0/groups/${groupId}/onenote/pages`) {
        return pageResponse.value;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { groupName: groupName } });
    assert(loggerLogSpy.calledWith(pageResponse.value));
  });

  it('lists Microsoft OneNote pages for site', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/sites/${siteId}/onenote/pages`) {
        return pageResponse.value;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts: any) => {
      const url = new URL(webUrl);
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${url.hostname}:${url.pathname}?$select=id`) {
        return { id: siteId };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl } });
    assert(loggerLogSpy.calledWith(pageResponse.value));
  });

  it('throws error when retrieving Microsoft OneNote notebooks for site and no site with specified url is found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const url = new URL(webUrl);
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/${url.hostname}:${url.pathname}?$select=id`) {
        throw {
          "error": {
            "code": "itemNotFound",
            "message": "Requested site could not be found",
            "innerError": {
              "date": "2023-01-07T11:55:48",
              "request-id": "18925839-f7e6-4827-bcb2-935a7836e734",
              "client-request-id": "18925839-f7e6-4827-bcb2-935a7836e734"
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl } } as any), new CommandError('Requested site could not be found'));
  });

  it('throws error if group by displayName returns no results', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(groupName)}'`) {
        return [];
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { groupName: groupName } } as any), new CommandError(`The specified group '${groupName}' does not exist.`));
  });

  it('throws an error if group by displayName returns multiple results', async () => {
    const duplicateGroupId = '9f3c2c36-1682-4922-9ae1-f57d2caf0de1';
    sinon.stub(odata, 'getAllItems').callsFake(async (url: string) => {
      if (url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(groupName)}'`) {
        return [{
          "id": groupId,
          "description": groupName,
          "displayName": groupName
        }, {
          "id": duplicateGroupId,
          "description": groupName,
          "displayName": groupName
        }];
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { groupName: groupName } } as any), new CommandError(`Multiple groups with name '${groupName}' found: ${groupId},${duplicateGroupId}.`));
  });
});
