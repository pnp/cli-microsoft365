import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./sitescript-add');

describe(commands.SITESCRIPT_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });


  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITESCRIPT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds new site script', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='Contoso'&@description='My%20contoso%20script'`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          abc: 'def'
        })) {
        return Promise.resolve({
          "Content": null,
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { title: 'Contoso', description: 'My contoso script', content: JSON.stringify({ "abc": "def" }) } });
    assert(loggerLogSpy.calledWith({
      "Content": null,
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 0
    }));
  });

  it('adds new site script (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='Contoso'&@description='My%20contoso%20script'`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          abc: 'def'
        })) {
        return Promise.resolve({
          "Content": null,
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, title: 'Contoso', description: 'My contoso script', content: JSON.stringify({ "abc": "def" }) } });
    assert(loggerLogSpy.calledWith({
      "Content": null,
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 0
    }));
  });

  it('doesn\'t fail when description not passed', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='Contoso'&@description=''`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          abc: 'def'
        })) {
        return Promise.resolve({
          "Content": null,
          "Description": "",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { title: 'Contoso', description: '', content: JSON.stringify({ "abc": "def" }) } });
    assert(loggerLogSpy.calledWith({
      "Content": null,
      "Description": "",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 0
    }));
  });

  it('escapes special characters in user input', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title, Description=@description)?@title='Contoso%20script'&@description='My%20contoso%20script'`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          abc: 'def'
        })) {
        return Promise.resolve({
          "Content": null,
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso script",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, title: 'Contoso script', description: 'My contoso script', content: JSON.stringify({ "abc": "def" }) } });
    assert(loggerLogSpy.calledWith({
      "Content": null,
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso script",
      "Version": 0
    }));
  });

  it('correctly handles OData error when creating site script', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    await assert.rejects(command.action(logger, { options: { title: 'Contoso', content: JSON.stringify({}) } } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying title', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--title') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying script content', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--content') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if script content is not a valid JSON string', async () => {
    const actual = await command.validate({ options: { title: 'Contoso', content: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when title specified and  script content is valid JSON', async () => {
    const actual = await command.validate({ options: { title: 'Contoso', content: JSON.stringify({}) } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});