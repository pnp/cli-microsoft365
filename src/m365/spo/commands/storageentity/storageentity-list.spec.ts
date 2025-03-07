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
import commands from '../../commands.js';
import command from './storageentity-list.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.STORAGEENTITY_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.STORAGEENTITY_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the list of configured tenant properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {
            storageentitiesindex: JSON.stringify({
              'Property1': {
                Value: 'dolor1'
              },
              'Property2': {
                Comment: 'Lorem2',
                Description: 'ipsum2',
                Value: 'dolor2'
              }
            })
          };
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    assert(loggerLogSpy.calledWith([
      {
        Key: 'Property1',
        Description: undefined,
        Comment: undefined,
        Value: 'dolor1'
      },
      {
        Key: 'Property2',
        Description: 'ipsum2',
        Comment: 'Lorem2',
        Value: 'dolor2'
      }
    ]));
  });

  it('doesn\'t fail if no tenant properties have been configured', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { storageentitiesindex: '' };
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
  });

  it('doesn\'t fail if tenant properties web property value is empty', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return {};
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    let correctResponse: boolean = false;
    log.forEach(l => {
      if (!l || typeof l !== 'string') {
        return;
      }

      if (l.indexOf('No tenant properties found') > -1) {
        correctResponse = true;
      }
    });
    assert(correctResponse, 'Incorrect response');
  });

  it('doesn\'t fail if tenant properties web property value is empty JSON object', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { storageentitiesindex: JSON.stringify({}) };
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
  });

  it('doesn\'t fail if tenant properties web property value is empty JSON object (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { storageentitiesindex: JSON.stringify({}) };
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    let correctResponse: boolean = false;
    log.forEach(l => {
      if (!l || typeof l !== 'string') {
        return;
      }

      if (l.indexOf('No tenant properties found') > -1) {
        correctResponse = true;
      }
    });
    assert(correctResponse, 'Incorrect response');
  });

  it('doesn\'t fail if tenant properties web property value is invalid JSON', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return { storageentitiesindex: 'a' };
        }
      }

      throw 'Invalid request';
    });

    let errorMessage;
    try {
      JSON.parse('a');
    }
    catch (err: any) {
      errorMessage = err.message;
    }

    await assert.rejects(command.action(logger, { options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } } as any), new CommandError(`${errorMessage}`));
  });

  it('requires app catalog URL', () => {
    const options = command.options;
    let requiresAppCatalogUrl = false;
    options.forEach(o => {
      if (o.option.indexOf('<appCatalogUrl>') > -1) {
        requiresAppCatalogUrl = true;
      }
    });
    assert(requiresAppCatalogUrl);
  });

  it('accepts valid SharePoint Online app catalog URL', async () => {
    const actual = await command.validate({ options: { appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts valid SharePoint Online site URL', async () => {
    const actual = await command.validate({ options: { appCatalogUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online URL', async () => {
    const url = 'http://contoso';
    const actual = await command.validate({ options: { appCatalogUrl: url } }, commandInfo);
    assert.strictEqual(actual, `'${url}' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation when no SharePoint Online app catalog URL specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, 'Required option appCatalogUrl not specified');
  });

  it('handles promise rejection', async () => {
    sinon.stub(request, 'get').rejects(new Error('error'));

    await assert.rejects(command.action(logger, { options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } } as any), new CommandError('error'));
  });
});
