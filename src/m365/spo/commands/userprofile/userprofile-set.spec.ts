import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./userprofile-set');

describe(commands.USERPROFILE_SET, () => {
  let log: any[];
  let logger: Logger;
  const spoUrl = 'https://contoso.sharepoint.com';

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
    auth.service.spoUrl = spoUrl;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USERPROFILE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates single valued profile property', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`${spoUrl}/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty`) > -1) {
        return {
          "odata.null": true
        };
      }
      throw 'Invalid request';
    });

    const data: any = {
      'accountName': `i:0#.f|membership|john.doe@mytenant.onmicrosoft.com`,
      'propertyName': 'SPS-JobTitle',
      'propertyValue': 'Senior Developer'
    };

    await command.action(logger, {
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-JobTitle',
        propertyValue: 'Senior Developer',
        debug: true
      }
    });
    const lastCall = postStub.lastCall.args[0];
    assert.strictEqual(JSON.stringify(lastCall.data), JSON.stringify(data));
  });

  it('updates multi valued profile property', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`${spoUrl}/_api/SP.UserProfiles.PeopleManager/SetMultiValuedProfileProperty`) > -1) {
        return {
          "odata.null": true
        };
      }
      throw 'Invalid request';
    });

    const data: any = {
      'accountName': `i:0#.f|membership|john.doe@mytenant.onmicrosoft.com`,
      'propertyName': 'SPS-Skills',
      'propertyValues': ['CSS', 'HTML']
    };

    await command.action(logger, {
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-Skills',
        propertyValue: 'CSS, HTML'
      }
    });
    const lastCall = postStub.lastCall.args[0];
    assert.strictEqual(JSON.stringify(lastCall.data), JSON.stringify(data));
  });

  it('correctly handles error while updating profile property', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        userName: 'john.doe@mytenant.onmicrosoft.com',
        propertyName: 'SPS-JobTitle',
        propertyValue: 'Senior Developer'
      }
    } as any), new CommandError('An error has occurred'));
  });
});
