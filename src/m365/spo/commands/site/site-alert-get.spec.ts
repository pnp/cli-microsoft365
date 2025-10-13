import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
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
import command from './site-alert-get.js';

describe(commands.SITE_ALERT_GET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/Marketing';
  const alertId = '7cbb4c8d-8e4d-4d2e-9c6f-3f1d8b2e6a0e';
  const alertResponse = {
    AlertFrequency: 0,
    AlertTemplateName: 'SPAlertTemplateType.DocumentLibrary',
    AlertType: 0,
    AlwaysNotify: false,
    DeliveryChannels: 1,
    EventType: -1,
    Filter: '',
    ID: 'a188ee89-72e2-4327-9802-8d0c408ec129',
    List: {
      Id: '1ec04825-b082-46f8-9a1c-b6b54d83bc46',
      RootFolder: {
        ServerRelativeUrl: '/sites/Marketing/Documents'
      },
      Title: 'Documents'
    },
    Properties: [
      {
        Key: 'dispformurl',
        Value: 'Documents/Forms/DispForm.aspx',
        ValueType: 'Edm.String'
      },
      {
        Key: 'filterindex',
        Value: '0',
        ValueType: 'Edm.String'
      },
      {
        Key: 'defaultitemopen',
        Value: 'Browser',
        ValueType: 'Edm.String'
      },
      {
        Key: 'sendurlinsms',
        Value: 'False',
        ValueType: 'Edm.String'
      },
      {
        Key: 'mobileurl',
        Value: 'https://contoso.sharepoint.com/_layouts/15/mobile/',
        ValueType: 'Edm.String'
      },
      {
        Key: 'eventtypeindex',
        Value: '0',
        ValueType: 'Edm.String'
      },
      {
        Key: 'siteurl',
        Value: 'https://contoso.sharepoint.com',
        ValueType: 'Edm.String'
      }
    ],
    Status: 0,
    Title: 'Documents',
    User: {
      Email: 'admin@contoso.onmicrosoft.com',
      Expiration: '',
      Id: 10,
      IsEmailAuthenticationGuestUser: false,
      IsHiddenInUI: false,
      IsShareByEmailGuestUser: false,
      IsSiteAdmin: true,
      LoginName: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
      PrincipalType: 1,
      Title: 'Admin User',
      UserId: {
        NameId: '100320009d8267fc',
        NameIdIssuer: 'urn:federation:microsoftonline'
      },
      UserPrincipalName: 'admin@contoso.onmicrosoft.com'
    },
    UserId: 10
  };

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    assert.strictEqual(command.name.startsWith(commands.SITE_ALERT_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves alert details by ID', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/Alerts/GetById('${alertId}')?$expand=List,User,List/Rootfolder&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl`) {
        return alertResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        id: alertId
      }
    });
    assert(loggerLogSpy.calledWith(alertResponse));
  });

  it('correctly handles error when alert does not exist', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-2146232832, Microsoft.SharePoint.SPException',
          message: {
            lang: 'en-US',
            value: 'The alert you are trying to access does not exist or has just been deleted.  '
          }
        }
      }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/Web/Alerts/GetById('${alertId}')?$expand=List,User,List/Rootfolder&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: {
          webUrl: webUrl,
          id: alertId
        }
      } as any),
      new CommandError(error.error['odata.error'].message.value)
    );
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'foo',
      id: alertId
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: webUrl,
      id: alertId
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: webUrl,
      id: 'invalid-guid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: webUrl,
      id: alertId
    });
    assert.strictEqual(actual.success, true);
  });
});

