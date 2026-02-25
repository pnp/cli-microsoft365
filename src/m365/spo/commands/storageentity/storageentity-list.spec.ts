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
import command, { options } from './storageentity-list.js';

describe(commands.STORAGEENTITY_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      spo.getTenantAppCatalogUrl
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
      if (opts.url === 'https://contoso.sharepoint.com/sites/appcatalog/_api/web/AllProperties?$select=storageentitiesindex') {
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

      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
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

  it('retrieves tenant properties using tenant app catalog URL when appCatalogUrl is not specified', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/appcatalog');

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/appcatalog/_api/web/AllProperties?$select=storageentitiesindex') {
        return {
          storageentitiesindex: JSON.stringify({
            'Property1': {
              Value: 'dolor1'
            }
          })
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert(loggerLogSpy.calledWith([
      {
        Key: 'Property1',
        Description: undefined,
        Comment: undefined,
        Value: 'dolor1'
      }
    ]));
  });

  it('throws error when tenant app catalog is not found and appCatalogUrl is not specified', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(null);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError('Tenant app catalog URL not found. Specify the URL of the app catalog site using the appCatalogUrl option.'));
  });

  it('doesn\'t fail if no tenant properties have been configured', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/appcatalog/_api/web/AllProperties?$select=storageentitiesindex') {
        return { storageentitiesindex: '' };
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(loggerLogSpy.calledWith([]));
  });

  it('doesn\'t fail if tenant properties web property value is empty', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/appcatalog/_api/web/AllProperties?$select=storageentitiesindex') {
        return {};
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(loggerLogSpy.calledWith([]));
  });

  it('doesn\'t fail if tenant properties web property value is empty JSON object', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/appcatalog/_api/web/AllProperties?$select=storageentitiesindex') {
        return { storageentitiesindex: JSON.stringify({}) };
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(loggerLogSpy.calledWith([]));
  });

  it('doesn\'t fail if tenant properties web property value is empty JSON object (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/appcatalog/_api/web/AllProperties?$select=storageentitiesindex') {
        return { storageentitiesindex: JSON.stringify({}) };
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(loggerLogSpy.calledWith([]));
  });

  it('doesn\'t fail if tenant properties web property value is invalid JSON', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/appcatalog/_api/web/AllProperties?$select=storageentitiesindex') {
        return { storageentitiesindex: 'a' };
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) }), new CommandError(`${errorMessage}`));
  });

  it('fails validation if appCatalogUrl is not a valid URL', () => {
    const actual = commandOptionsSchema.safeParse({ appCatalogUrl: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when appCatalogUrl is a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when appCatalogUrl is not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('handles promise rejection', async () => {
    sinon.stub(request, 'get').rejects(new Error('error'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) }), new CommandError('error'));
  });
});
