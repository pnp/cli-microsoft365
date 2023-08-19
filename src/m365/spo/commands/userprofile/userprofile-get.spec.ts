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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './userprofile-get.js';

describe(commands.USERPROFILE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = true;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USERPROFILE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets userprofile information about the specified user', async () => {
    const profile = {
      "AccountName": "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com",
      "DirectReports": [],
      "DisplayName": "Dipen Shah",
      "Email": "dips1802@dev1802.onmicrosoft.com",
      "ExtendedManagers": [],
      "ExtendedReports": [
        "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com"
      ],
      "IsFollowed": false,
      "LatestPost": null,
      "Peers": [],
      "PersonalSiteHostUrl": "https://contoso-my.sharepoint.com:443/",
      "PersonalUrl": "https://contoso-my.sharepoint.com/personal/dips1802_dev1802_onmicrosoft_com/",
      "PictureUrl": null,
      "Title": null,
      "UserProfileProperties": [
        {
          "Key": "UserProfile_GUID",
          "Value": "f3f102bb-7ac7-408e-9184-384062abd0d5"
        },
        {
          "Key": "SID",
          "Value": "i:0h.f|membership|10032000840f3681@live.com"
        }]
    };
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor') > -1) {
        // we need to clone the object because it's changed in the command
        // and would skew the comparison in the test outcome
        return JSON.parse(JSON.stringify(profile));
      }
      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        output: 'text',
        userName: 'john.doe@contoso.onmicrosoft.com'
      }
    } as any);
    const loggedProfile = JSON.parse(JSON.stringify(profile));
    loggedProfile.UserProfileProperties = JSON.stringify(loggedProfile.UserProfileProperties);
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify(loggedProfile));
  });

  it('gets userprofile information about the specified user output json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor') > -1) {
        return {
          "AccountName": "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com",
          "DirectReports": [],
          "DisplayName": "Dipen Shah",
          "Email": "dips1802@dev1802.onmicrosoft.com",
          "ExtendedManagers": [],
          "ExtendedReports": [
            "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com"
          ],
          "IsFollowed": false,
          "LatestPost": null,
          "Peers": [],
          "PersonalSiteHostUrl": "https://contoso-my.sharepoint.com:443/",
          "PersonalUrl": "https://contoso-my.sharepoint.com/personal/dips1802_dev1802_onmicrosoft_com/",
          "PictureUrl": null,
          "Title": null
        };
      }
      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        userName: 'john.doe@contoso.onmicrosoft.com'
      }
    } as any);
    assert(loggerLogSpy.calledWith({
      "AccountName": "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com",
      "DirectReports": [],
      "DisplayName": "Dipen Shah",
      "Email": "dips1802@dev1802.onmicrosoft.com",
      "ExtendedManagers": [],
      "ExtendedReports": [
        "i:0#.f|membership|dips1802@dev1802.onmicrosoft.com"
      ],
      "IsFollowed": false,
      "LatestPost": null,
      "Peers": [],
      "PersonalSiteHostUrl": "https://contoso-my.sharepoint.com:443/",
      "PersonalUrl": "https://contoso-my.sharepoint.com/personal/dips1802_dev1802_onmicrosoft_com/",
      "PictureUrl": null,
      "Title": null
    }));
  });

  it('supports specifying userName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--userName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        userName: 'john.doe@contoso.onmicrosoft.com'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if the user principal name is not a valid', async () => {
    const actual = await command.validate({ options: { userName: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the user principal name is a valid', async () => {
    const actual = await command.validate({ options: { userName: 'john.doe@mytenant.onmicrosoft.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
