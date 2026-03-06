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
import command, { options } from './storageentity-get.js';

describe(commands.STORAGEENTITY_GET, () => {
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
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetStorageEntity('existingproperty')`) {
        return { Comment: 'Lorem', Description: 'ipsum', Value: 'dolor' };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetStorageEntity('propertywithoutdescription')`) {
        return { Comment: 'Lorem', Value: 'dolor' };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetStorageEntity('propertywithoutcomments')`) {
        return { Description: 'ipsum', Value: 'dolor' };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetStorageEntity('nonexistingproperty')`) {
        return { "odata.null": true };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetStorageEntity('%23myprop')`) {
        return { Description: 'ipsum', Value: 'dolor' };
      }

      throw 'Invalid request';
    });
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
      spo.getTenantAppCatalogUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.STORAGEENTITY_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the details of an existing tenant property', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: 'existingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(loggerLogSpy.calledWith({
      Key: 'existingproperty',
      Value: 'dolor',
      Description: 'ipsum',
      Comment: 'Lorem'
    }));
  });

  it('retrieves the details of an existing tenant property without a description', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: 'propertywithoutdescription', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(loggerLogSpy.calledWith({
      Key: 'propertywithoutdescription',
      Value: 'dolor',
      Description: undefined,
      Comment: 'Lorem'
    }));
  });

  it('retrieves the details of an existing tenant property without a comment', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ key: 'propertywithoutcomments', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(loggerLogSpy.calledWith({
      Key: 'propertywithoutcomments',
      Value: 'dolor',
      Description: 'ipsum',
      Comment: undefined
    }));
  });

  it('retrieves tenant property using tenant app catalog URL when appCatalogUrl is not specified', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/appcatalog');

    await command.action(logger, { options: commandOptionsSchema.parse({ key: 'existingproperty' }) });
    assert(loggerLogSpy.calledWith({
      Key: 'existingproperty',
      Value: 'dolor',
      Description: 'ipsum',
      Comment: 'Lorem'
    }));
  });

  it('throws error when tenant app catalog is not found and appCatalogUrl is not specified', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(null);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ key: 'existingproperty' }) }),
      new CommandError('Tenant app catalog URL not found. Specify the URL of the app catalog site using the appCatalogUrl option.'));
  });

  it('handles a non-existent tenant property', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ key: 'nonexistingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
  });

  it('handles a non-existent tenant property (debug)', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: 'nonexistingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    let correctValue: boolean = false;
    log.forEach(l => {
      if (l &&
        typeof l === 'string' &&
        l.includes('Property with key nonexistingproperty not found')) {
        correctValue = true;
      }
    });
    assert(correctValue);
  });

  it('escapes special characters in property name', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: '#myprop', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(loggerLogSpy.calledWith({
      Key: '#myprop',
      Value: 'dolor',
      Description: 'ipsum',
      Comment: undefined
    }));
  });

  it('fails validation if appCatalogUrl is not a valid URL', () => {
    const actual = commandOptionsSchema.safeParse({ key: 'prop', appCatalogUrl: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when appCatalogUrl is a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ key: 'prop', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when appCatalogUrl is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ key: 'prop' });
    assert.strictEqual(actual.success, true);
  });

  it('handles promise rejection', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('error'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: '#myprop', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) }), new CommandError('error'));
  });
});
