import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { telemetry } from '../../../../telemetry.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { z } from 'zod';
import commands from '../../commands.js';
import command from './site-alert-list.js';

describe(commands.SITE_ALERT_LIST, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let loggerLogSpy: sinon.SinonSpy;

  const webUrl = 'https://contoso.sharepoint.com/sites/marketing';
  const listId = '39d9e102-9e8f-4e74-8f17-84a92f972fcf';
  const listTitle = 'Tasks';
  const listUrl = '/sites/marketing/lists/tasks';
  const userName = 'jane.doe@contoso.com';
  const userId = '7cbb4c8d-8e4d-4d2e-9c6f-3f1d8b2e6a0e';

  const alertResponse = [
    {
      "AlertFrequency": 0,
      "AlertTemplateName": "SPAlertTemplateType.DocumentLibrary",
      "AlertType": 0,
      "AlwaysNotify": false,
      "DeliveryChannels": 1,
      "EventType": -1,
      "Filter": "",
      "ID": "edf49cf4-d91e-4084-866d-2b9ff5582189",
      "Properties": [
        {
          "Key": "dispformurl",
          "Value": "Tasks/Forms/DispForm.aspx",
          "ValueType": "Edm.String"
        },
        {
          "Key": "filterindex",
          "Value": "0",
          "ValueType": "Edm.String"
        },
        {
          "Key": "defaultitemopen",
          "Value": "Browser",
          "ValueType": "Edm.String"
        },
        {
          "Key": "sendurlinsms",
          "Value": "False",
          "ValueType": "Edm.String"
        },
        {
          "Key": "mobileurl",
          "Value": "https://contoso.sharepoint.com/_layouts/15/mobile/",
          "ValueType": "Edm.String"
        },
        {
          "Key": "eventtypeindex",
          "Value": "0",
          "ValueType": "Edm.String"
        },
        {
          "Key": "webUrl",
          "Value": "https://contoso.sharepoint.com",
          "ValueType": "Edm.String"
        }
      ],
      "Status": 0,
      "Title": listTitle,
      "UserId": 8,
      "List": {
        "Id": listId,
        "Title": listTitle,
        "RootFolder": {
          "ServerRelativeUrl": listUrl
        }
      },
      "User": {
        "Id": 8,
        "IsHiddenInUI": false,
        "LoginName": `i:0#.f|membership|${userName}`,
        "Title": "Jane Doe",
        "PrincipalType": 1,
        "Email": userName,
        "Expiration": "",
        "IsEmailAuthenticationGuestUser": false,
        "IsShareByEmailGuestUser": false,
        "IsSiteAdmin": false,
        "UserId": {
          "NameId": "f7e8d9c2b1a5e3d6",
          "NameIdIssuer": "urn:federation:microsoftonline"
        },
        "UserPrincipalName": userName
      }
    }
  ];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    auth.connection.active = true;
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
      odata.getAllItems,
      entraUser.getUpnByUserId
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_ALERT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if webUrl is not a valid URL', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if listId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both listId and listUrl are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId, listUrl: listUrl });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both listUrl and listTitle are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listUrl: listUrl, listTitle: listTitle });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, userId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if userName is not a valid user principal name', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, userName: 'invalid' });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, userId: userId, userName: userName });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when only webUrl is specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when only webUrl and listId are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listId: listId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when only webUrl and userId are specified', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, userId: userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when all parameters are valid', async () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: webUrl, listTitle: listTitle, userName: userName });
    assert.strictEqual(actual.success, true);
  });

  it('successfully gets all alerts', async () => {
    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${webUrl}/_api/web/alerts?$expand=List,User,List/Rootfolder,Item&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl,Item/ID,Item/FileRef,Item/Guid`) {
        return alertResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, verbose: true } });
    assert(odataStub.calledOnce);
    assert(loggerLogSpy.calledWith(alertResponse));
  });

  it('successfully gets all alerts when listId is specified', async () => {
    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${webUrl}/_api/web/alerts?$expand=List,User,List/Rootfolder,Item&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl,Item/ID,Item/FileRef,Item/Guid&$filter=List/Id eq guid'${listId}'`) {
        return alertResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listId: listId, verbose: true } });
    assert(odataStub.calledOnce);
    assert(loggerLogSpy.calledWith(alertResponse));
  });

  it('successfully gets all alerts when listUrl is specified', async () => {
    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${webUrl}/_api/web/alerts?$expand=List,User,List/Rootfolder,Item&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl,Item/ID,Item/FileRef,Item/Guid&$filter=List/RootFolder/ServerRelativeUrl eq '${formatting.encodeQueryParameter(listUrl)}'`) {
        return alertResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listUrl: listUrl, verbose: true } });
    assert(odataStub.calledOnce);
    assert(loggerLogSpy.calledWith(alertResponse));
  });

  it('successfully gets all alerts when listTitle is specified', async () => {
    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${webUrl}/_api/web/alerts?$expand=List,User,List/Rootfolder,Item&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl,Item/ID,Item/FileRef,Item/Guid&$filter=List/Title eq '${listTitle}'`) {
        return alertResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listTitle: listTitle, verbose: true } });
    assert(odataStub.calledOnce);
    assert(loggerLogSpy.calledWith(alertResponse));
  });

  it('successfully gets all alerts when userName is specified', async () => {
    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${webUrl}/_api/web/alerts?$expand=List,User,List/Rootfolder,Item&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl,Item/ID,Item/FileRef,Item/Guid&$filter=User/UserPrincipalName eq '${formatting.encodeQueryParameter(userName)}'`) {
        return alertResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, userName: userName, verbose: true } });
    assert(odataStub.calledOnce);
    assert(loggerLogSpy.calledWith(alertResponse));
  });

  it('successfully gets all alerts when listId and userId are specified', async () => {
    sinon.stub(entraUser, 'getUpnByUserId').resolves(userName);

    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `${webUrl}/_api/web/alerts?$expand=List,User,List/Rootfolder,Item&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl,Item/ID,Item/FileRef,Item/Guid&$filter=List/Id eq guid'${listId}' and User/UserPrincipalName eq '${formatting.encodeQueryParameter(userName)}'`) {
        return alertResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, listId: listId, userId: userId, verbose: true } });
    assert(odataStub.calledOnce);
    assert(loggerLogSpy.calledWith(alertResponse));
  });

  it('correctly handles error when retrieving alerts', async () => {
    const error = {
      error: {
        code: 'UnknownError',
        message: 'An unknown error has occurred.'
      }
    };
    sinon.stub(odata, 'getAllItems').rejects(error);

    await assert.rejects(command.action(logger, { options: {} }),
      new CommandError(`An unknown error has occurred.`));
  });
});