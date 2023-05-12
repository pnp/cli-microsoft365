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
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./sitedesign-rights-grant');

describe(commands.SITEDESIGN_RIGHTS_GRANT, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITEDESIGN_RIGHTS_GRANT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('grants rights on the specified site design to the specified principal', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          "id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
          "principalNames": ["PattiF"],
          "grantedRights": "1"
        })) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { siteDesignId: '9b142c22-037f-4a7f-9017-e9d8c0e34b98', principals: 'PattiF', rights: 'View' } });
    assert(loggerLogSpy.notCalled);
  });

  it('grants rights on the specified site design to the specified principals', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          "id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
          "principalNames": ["PattiF", "AdeleV"],
          "grantedRights": "1"
        })) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { siteDesignId: '9b142c22-037f-4a7f-9017-e9d8c0e34b98', principals: 'PattiF,AdeleV', rights: 'View' } });
    assert(loggerLogSpy.notCalled);
  });

  it('grants rights on the specified site design to the specified principals (email)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          "id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
          "principalNames": ["PattiF@contoso.com", "AdeleV@contoso.com"],
          "grantedRights": "1"
        })) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { siteDesignId: '9b142c22-037f-4a7f-9017-e9d8c0e34b98', principals: 'PattiF@contoso.com,AdeleV@contoso.com', rights: 'View' } });
  });

  it('grants rights on the specified site design to the specified principals separated with an extra space', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          "id": "9b142c22-037f-4a7f-9017-e9d8c0e34b98",
          "principalNames": ["PattiF", "AdeleV"],
          "grantedRights": "1"
        })) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { siteDesignId: '9b142c22-037f-4a7f-9017-e9d8c0e34b98', principals: 'PattiF, AdeleV', rights: 'View' } });
  });

  it('correctly handles OData error when granting rights', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, {
      options: {
        siteDesignId: '9b142c22-037f-4a7f-9017-e9d8c0e34b98',
        principals: 'PattiF',
        rights: 'View'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteDesignId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying principals', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--principals') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying rights', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--rights') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { siteDesignId: 'abc', principals: 'PattiF', rights: 'View' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified rights value is invalid', async () => {
    const actual = await command.validate({ options: { siteDesignId: '9b142c22-037f-4a7f-9017-e9d8c0e34b98', principals: 'PattiF', rights: 'Invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required parameters are valid', async () => {
    const actual = await command.validate({ options: { siteDesignId: '9b142c22-037f-4a7f-9017-e9d8c0e34b98', principals: 'PattiF', rights: 'View' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if all required parameters are valid (multiple principals)', async () => {
    const actual = await command.validate({ options: { siteDesignId: '9b142c22-037f-4a7f-9017-e9d8c0e34b98', principals: 'PattiF,AdeleV', rights: 'View' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
