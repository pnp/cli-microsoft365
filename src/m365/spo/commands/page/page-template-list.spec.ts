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
import commands from '../../commands';
import { templatesMock } from './page-template-list.mock';
const command: Command = require('./page-template-list');

describe(commands.PAGE_TEMPLATE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_TEMPLATE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Title', 'FileName', 'Id', 'PageLayoutType', 'Url']);
  });

  it('list all page templates', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/templates`) > -1) {
        return Promise.resolve(templatesMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a' } });
    assert(loggerLogSpy.calledWith([...templatesMock.value]));
  });

  it('list all page templates (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/templates`) > -1) {
        return Promise.resolve(templatesMock);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } });
    assert(loggerLogSpy.calledWith([...templatesMock.value]));
  });

  it('correctly handles no page templates', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/templates`) > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles OData error when retrieving page templates', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    await assert.rejects(command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a' } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles error when retrieving page templates on a site which does not have any', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({ response: { status: 404 } });
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/team-a' } } as any);
    assert(loggerLogSpy.calledWith([]));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the webUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
