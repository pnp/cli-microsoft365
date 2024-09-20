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
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './sitescript-get.js';

describe(commands.SITESCRIPT_GET, () => {
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
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITESCRIPT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about the specified site script', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata') {
        return {
          "Content": JSON.stringify({
            "$schema": "schema.json",
            "actions": [
              {
                "verb": "applyTheme",
                "themeName": "Contoso Theme"
              }
            ],
            "bindata": {},
            "version": 1
          }),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } });
    assert(loggerLogSpy.calledOnceWithExactly({
      "Content": {
        "$schema": "schema.json",
        "actions": [
          {
            "verb": "applyTheme",
            "themeName": "Contoso Theme"
          }
        ],
        "bindata": {},
        "version": 1
      },
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 1
    }));
  });

  it('gets information about the specified site script (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata') {
        return {
          "Content": JSON.stringify({
            "$schema": "schema.json",
            "actions": [
              {
                "verb": "applyTheme",
                "themeName": "Contoso Theme"
              }
            ],
            "bindata": {},
            "version": 1
          }),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } });
    assert(loggerLogSpy.calledOnceWithExactly({
      "Content": {
        "$schema": "schema.json",
        "actions": [
          {
            "verb": "applyTheme",
            "themeName": "Contoso Theme"
          }
        ],
        "bindata": {},
        "version": 1
      },
      "Description": "My contoso script",
      "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
      "Title": "Contoso",
      "Version": 1
    }));
  });

  it('gets the specified site script contents', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata') {
        return {
          "Content": JSON.stringify({
            "$schema": "schema.json",
            "actions": [
              {
                "verb": "applyTheme",
                "themeName": "Contoso Theme"
              }
            ],
            "bindata": {},
            "version": 1
          }),
          "Description": "My contoso script",
          "Id": "0f27a016-d277-4bb4-b3c3-b5b040c9559b",
          "Title": "Contoso",
          "Version": 1
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b', content: true } });
    assert(loggerLogSpy.calledOnceWithExactly({
      "$schema": "schema.json",
      "actions": [
        {
          "verb": "applyTheme",
          "themeName": "Contoso Theme"
        }
      ],
      "bindata": {},
      "version": 1
    }));
  });

  it('correctly handles error when site script not found', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });

    await assert.rejects(command.action(logger, { options: { id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } } as any), new CommandError('File Not Found.'));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
