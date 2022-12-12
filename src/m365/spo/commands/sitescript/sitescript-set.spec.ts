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
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./sitescript-set');

describe(commands.SITESCRIPT_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITESCRIPT_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates title of an existing site script', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            'Id': '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            'Title': 'Contoso'
          }
        })) {
        return Promise.resolve({
          "Content": JSON.stringify({}),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', title: 'Contoso' } });
    assert(loggerLogSpy.calledWith({
      "Content": JSON.stringify({}),
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 0
    }));
  });

  it('updates title of an existing site script (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            'Id': '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            'Title': 'Contoso'
          }
        })) {
        return Promise.resolve({
          "Content": JSON.stringify({}),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', title: 'Contoso' } });
    assert(loggerLogSpy.calledWith({
      "Content": JSON.stringify({}),
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 0
    }));
  });

  it('updates description of an existing site script', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            'Id': '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            'Description': 'My contoso script'
          }
        })) {
        return Promise.resolve({
          "Content": JSON.stringify({}),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 0
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', description: 'My contoso script' } });
    assert(loggerLogSpy.calledWith({
      "Content": JSON.stringify({}),
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 0
    }));
  });

  it('updates version of an existing site script', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            'Id': '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            'Version': 1
          }
        })) {
        return Promise.resolve({
          "Content": JSON.stringify({}),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', version: '1' } });
    assert(loggerLogSpy.calledWith({
      "Content": JSON.stringify({}),
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 1
    }));
  });

  it('updates content of an existing site script', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            'Id': '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            'Content': JSON.stringify({})
          }
        })) {
        return Promise.resolve({
          "Content": JSON.stringify({}),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', content: JSON.stringify({}) } });
    assert(loggerLogSpy.calledWith({
      "Content": JSON.stringify({}),
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 1
    }));
  });

  it('updates all properties of an existing site script', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`) > -1 &&
        JSON.stringify(opts.data) === JSON.stringify({
          updateInfo: {
            Id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b',
            Title: 'Contoso',
            Description: 'My contoso script',
            Version: 1,
            Content: JSON.stringify({})
          }
        })) {
        return Promise.resolve({
          "Content": JSON.stringify({}),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', title: 'Contoso', description: 'My contoso script', version: '1', content: JSON.stringify({}) } });
    assert(loggerLogSpy.calledWith({
      "Content": JSON.stringify({}),
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 1
    }));
  });

  it('correctly handles OData error when creating site script', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    await assert.rejects(command.action(logger, { options: { id: '449c0c6d-5380-4df2-b84b-622e0ac8ec24', title: 'Contoso' } } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
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

  it('supports specifying version', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--version') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if version is not a valid number', async () => {
    const actual = await command.validate({ options: { id: '449c0c6d-5380-4df2-b84b-622e0ac8ec24', version: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if script content is not a valid JSON string', async () => {
    const actual = await command.validate({ options: { id: '449c0c6d-5380-4df2-b84b-622e0ac8ec24', content: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id specified and valid GUID', async () => {
    const actual = await command.validate({ options: { id: '449c0c6d-5380-4df2-b84b-622e0ac8ec24' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when id and version specified and version is a number', async () => {
    const actual = await command.validate({ options: { id: '449c0c6d-5380-4df2-b84b-622e0ac8ec24', version: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when id and content specified and content is valid JSON', async () => {
    const actual = await command.validate({ options: { id: '449c0c6d-5380-4df2-b84b-622e0ac8ec24', content: JSON.stringify({}) } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
