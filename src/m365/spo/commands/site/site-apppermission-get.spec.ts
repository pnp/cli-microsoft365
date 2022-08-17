import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./site-apppermission-get');

describe(commands.SITE_APPPERMISSION_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      global.setTimeout
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
    assert.strictEqual(command.name.startsWith(commands.SITE_APPPERMISSION_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with an incorrect URL', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https;//contoso,sharepoint:com/sites/sitecollection-name',
        permissionId: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation with a correct URL', async () => {
    const actual = await command.validate({
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permissionId: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with a correct URL and a filter value', async () => {
    const actual = await command.validate({
      options: {
        permissionId: 'aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0',
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('returns a specific application permissions for the site', (done) => {
    const site = {
      "id": "contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000",
      "displayName": "OneDrive Team Site",
      "name": "1drvteam",
      "createdDateTime": "2017-05-09T20:56:00Z",
      "lastModifiedDateTime": "2017-05-09T20:56:01Z",
      "webUrl": "https://contoso.sharepoint.com/teams/1drvteam"
    };

    const response = {
      "id": "aTowaS50fG1zLnNwLmV4dHxmYzE1MzRlNy0yNTlkLTQ4MmEtODY4OC1kNmEzM2Q5YTBhMmNAZWUyYjdjMGMtZDI1My00YjI3LTk0NmItMDYzZGM4OWNlOGMy",
      "roles": [
        "write"
      ],
      "grantedToIdentities": [
        {
          "application": {
            "displayName": "Selected",
            "id": "fc1534e7-259d-482a-8688-d6a33d9a0a2c"
          }
        }
      ]
    };

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf(":/sites/sitecollection-name") > - 1) {
          return Promise.resolve(site);
        }
        return Promise.reject('Invalid request');
      });

    getRequestStub.onCall(1)
      .callsFake((opts) => {
        if ((opts.url as string).indexOf("contoso.sharepoint.com,00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000/permissions/") > - 1) {
          return Promise.resolve(response);
        }
        return Promise.reject('Invalid request');
      });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name',
        permissionId: 'aTowaS50fG1zLnNwLmV4dHxmYzE1MzRlNy0yNTlkLTQ4MmEtODY4OC1kNmEzM2Q5YTBhMmNAZWUyYjdjMGMtZDI1My00YjI3LTk0NmItMDYzZGM4OWNlOGMy',
        output: 'json'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([{
          appDisplayName: 'Selected',
          appId: 'fc1534e7-259d-482a-8688-d6a33d9a0a2c',
          permissionId: 'aTowaS50fG1zLnNwLmV4dHxmYzE1MzRlNy0yNTlkLTQ4MmEtODY4OC1kNmEzM2Q5YTBhMmNAZWUyYjdjMGMtZDI1My00YjI3LTk0NmItMDYzZGM4OWNlOGMy',
          roles: "write"
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when passing a site that does not exist', (done) => {
    const siteError = {
      "error": {
        "code": "itemNotFound",
        "message": "Requested site could not be found",
        "innerError": {
          "date": "2021-03-03T08:58:02",
          "request-id": "4e054f93-0eba-4743-be47-ce36b5f91120",
          "client-request-id": "dbd35b28-0ec3-6496-1279-0e1da3d028fe"
        }
      }
    };
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('non-existing') === -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject(siteError);
    });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/sitecollection-name-non-existing'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Requested site could not be found")));
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
