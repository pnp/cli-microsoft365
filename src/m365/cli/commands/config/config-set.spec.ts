import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { settingsNames } from '../../../../settingsNames.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './config-set.js';

describe(commands.CONFIG_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
  });

  afterEach(() => {
    sinonUtil.restore(cli.getConfig().set);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONFIG_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`sets ${settingsNames.showHelpOnFailure} property`, async () => {
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.showHelpOnFailure, value: false } });
    assert.strictEqual(actualKey, settingsNames.showHelpOnFailure, 'Invalid key');
    assert.strictEqual(actualValue, false, 'Invalid value');
  });

  it(`sets ${settingsNames.autoOpenLinksInBrowser} property`, async () => {
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.autoOpenLinksInBrowser, value: false } });
    assert.strictEqual(actualKey, settingsNames.autoOpenLinksInBrowser, 'Invalid key');
    assert.strictEqual(actualValue, false, 'Invalid value');
  });

  it(`sets ${settingsNames.output} property to 'text'`, async () => {
    const output = "text";
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);

    await command.action(logger, { options: { key: settingsNames.output, value: output } });
    assert.strictEqual(actualKey, settingsNames.output, 'Invalid key');
    assert.strictEqual(actualValue, output, 'Invalid value');
  });

  it(`sets ${settingsNames.output} property to 'json'`, async () => {
    const output = "json";
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);

    await command.action(logger, { options: { key: settingsNames.output, value: output } });
    assert.strictEqual(actualKey, settingsNames.output, 'Invalid key');
    assert.strictEqual(actualValue, output, 'Invalid value');
  });

  it(`sets ${settingsNames.output} property to 'csv'`, async () => {
    const output = "csv";
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);

    await command.action(logger, { options: { key: settingsNames.output, value: output } });
    assert.strictEqual(actualKey, settingsNames.output, 'Invalid key');
    assert.strictEqual(actualValue, output, 'Invalid value');
  });

  it(`sets ${settingsNames.csvHeader} property`, async () => {
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.csvHeader, value: false } });
    assert.strictEqual(actualKey, settingsNames.csvHeader, 'Invalid key');
    assert.strictEqual(actualValue, false, 'Invalid value');
  });

  it(`sets ${settingsNames.csvQuoted} property`, async () => {
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.csvQuoted, value: false } });
    assert.strictEqual(actualKey, settingsNames.csvQuoted, 'Invalid key');
    assert.strictEqual(actualValue, false, 'Invalid value');
  });

  it(`sets ${settingsNames.csvQuotedEmpty} property`, async () => {
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.csvQuotedEmpty, value: false } });
    assert.strictEqual(actualKey, settingsNames.csvQuotedEmpty, 'Invalid key');
    assert.strictEqual(actualValue, false, 'Invalid value');
  });

  it(`sets ${settingsNames.prompt} property`, async () => {
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.prompt, value: false } });
    assert.strictEqual(actualKey, settingsNames.prompt, 'Invalid key');
    assert.strictEqual(actualValue, false, 'Invalid value');
  });

  it(`sets ${settingsNames.authType} property`, async () => {
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.authType, value: 'deviceCode' } });
    assert.strictEqual(actualKey, settingsNames.authType, 'Invalid key');
    assert.strictEqual(actualValue, 'deviceCode', 'Invalid value');
  });

  it(`sets ${settingsNames.promptListPageSize} property`, async () => {
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.promptListPageSize, value: 10 } });
    assert.strictEqual(actualKey, settingsNames.promptListPageSize, 'Invalid key');
    assert.strictEqual(actualValue, 10, 'Invalid value');
  });

  it(`sets ${settingsNames.helpTarget} property`, async () => {
    const config = cli.getConfig();
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.helpTarget, value: 'console' } });
    assert.strictEqual(actualKey, settingsNames.helpTarget, 'Invalid key');
    assert.strictEqual(actualValue, 'console', 'Invalid value');
  });


  it('supports specifying key and value', () => {
    const options = command.options;
    let containsOptionKey = false;
    let containsOptionValue = false;
    options.forEach(o => {
      if (o.option.indexOf('--key') > -1) {
        containsOptionKey = true;
      }

      if (o.option.indexOf('--value') > -1) {
        containsOptionValue = true;
      }
    });
    assert(containsOptionKey && containsOptionValue);
  });

  it('fails validation if specified key is invalid ', async () => {
    const actual = await command.validate({ options: { key: 'invalid', value: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if setting is set to ${settingsNames.showHelpOnFailure} and value to true`, async () => {
    const actual = await command.validate({ options: { key: settingsNames.showHelpOnFailure, value: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it(`passes validation if setting is set to ${settingsNames.showHelpOnFailure} and value to false`, async () => {
    const actual = await command.validate({ options: { key: settingsNames.showHelpOnFailure, value: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified output type is invalid', async () => {
    const actual = await command.validate({ options: { key: settingsNames.output, value: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation for output type text', async () => {
    const actual = await command.validate({ options: { key: settingsNames.output, value: 'text' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for output type json', async () => {
    const actual = await command.validate({ options: { key: settingsNames.output, value: 'json' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for output type csv', async () => {
    const actual = await command.validate({ options: { key: settingsNames.output, value: 'csv' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified authType is invalid', async () => {
    const actual = await command.validate({ options: { key: settingsNames.authType, value: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation for authType type deviceCode', async () => {
    const actual = await command.validate({ options: { key: settingsNames.authType, value: 'deviceCode' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for authType type browser', async () => {
    const actual = await command.validate({ options: { key: settingsNames.authType, value: 'browser' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for authType type certificate', async () => {
    const actual = await command.validate({ options: { key: settingsNames.authType, value: 'certificate' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for authType type password', async () => {
    const actual = await command.validate({ options: { key: settingsNames.authType, value: 'password' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for authType type identity', async () => {
    const actual = await command.validate({ options: { key: settingsNames.authType, value: 'identity' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for authType type secret', async () => {
    const actual = await command.validate({ options: { key: settingsNames.authType, value: 'secret' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified error output type is invalid', async () => {
    const actual = await command.validate({ options: { key: settingsNames.errorOutput, value: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation for error output stdout', async () => {
    const actual = await command.validate({ options: { key: settingsNames.errorOutput, value: 'stdout' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for error output stderr', async () => {
    const actual = await command.validate({ options: { key: settingsNames.errorOutput, value: 'stderr' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified help mode is invalid', async () => {
    const actual = await command.validate({ options: { key: settingsNames.helpMode, value: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation for help mode options', async () => {
    const actual = await command.validate({ options: { key: settingsNames.helpMode, value: 'options' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for help mode examples', async () => {
    const actual = await command.validate({ options: { key: settingsNames.helpMode, value: 'examples' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for help mode remarks', async () => {
    const actual = await command.validate({ options: { key: settingsNames.helpMode, value: 'remarks' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for help mode response', async () => {
    const actual = await command.validate({ options: { key: settingsNames.helpMode, value: 'response' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for help mode full', async () => {
    const actual = await command.validate({ options: { key: settingsNames.helpMode, value: 'full' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified promptListPageSize value is a string', async () => {
    const actual = await command.validate({ options: { key: settingsNames.promptListPageSize, value: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified promptListPageSize value is 0', async () => {
    const actual = await command.validate({ options: { key: settingsNames.promptListPageSize, value: 0 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified promptListPageSize value is negative', async () => {
    const actual = await command.validate({ options: { key: settingsNames.promptListPageSize, value: -1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation for number value in promptListPageSize', async () => {
    const actual = await command.validate({ options: { key: settingsNames.promptListPageSize, value: 10 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified help target is invalid', async () => {
    const actual = await command.validate({ options: { key: settingsNames.helpTarget, value: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation for help target web', async () => {
    const actual = await command.validate({ options: { key: settingsNames.helpTarget, value: 'web' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for help target console', async () => {
    const actual = await command.validate({ options: { key: settingsNames.helpTarget, value: 'console' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified clientId is not a GUID', async () => {
    const actual = await command.validate({ options: { key: settingsNames.clientId, value: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if specified clientId is a GUID', async () => {
    const actual = await command.validate({ options: { key: settingsNames.clientId, value: '00000000-0000-0000-c000-000000000001' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if specified tenantId is not a GUID or common', async () => {
    const actual = await command.validate({ options: { key: settingsNames.tenantId, value: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if specified tenantId is a GUID', async () => {
    const actual = await command.validate({ options: { key: settingsNames.tenantId, value: '00000000-0000-0000-c000-000000000001' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if specified tenantId is common', async () => {
    const actual = await command.validate({ options: { key: settingsNames.tenantId, value: 'common' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
