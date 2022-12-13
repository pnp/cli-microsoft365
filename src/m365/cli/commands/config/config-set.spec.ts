import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import { settingsNames } from '../../../../settingsNames';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./config-set');

describe(commands.CONFIG_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    commandInfo = Cli.getCommandInfo(command);
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore(Cli.getInstance().config.set);
  });

  after(() => {
    sinonUtil.restore([
      telemetry.trackEvent,
      pid.getProcessName
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONFIG_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`sets ${settingsNames.showHelpOnFailure} property`, async () => {
    const config = Cli.getInstance().config;
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
    const config = Cli.getInstance().config;
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
    const config = Cli.getInstance().config;
    let actualKey: string = '', actualValue: any;
    sinon.restore();
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
    const config = Cli.getInstance().config;
    let actualKey: string = '', actualValue: any;
    sinon.restore();
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
    const config = Cli.getInstance().config;
    let actualKey: string = '', actualValue: any;
    sinon.restore();
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);

    await command.action(logger, { options: { key: settingsNames.output, value: output } });
    assert.strictEqual(actualKey, settingsNames.output, 'Invalid key');
    assert.strictEqual(actualValue, output, 'Invalid value');
  });

  it(`sets ${settingsNames.csvHeader} property`, async () => {
    const config = Cli.getInstance().config;
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
    const config = Cli.getInstance().config;
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
    const config = Cli.getInstance().config;
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
    const config = Cli.getInstance().config;
    let actualKey: string = '', actualValue: any;
    sinon.stub(config, 'set').callsFake(((key: string, value: any) => {
      actualKey = key;
      actualValue = value;
    }) as any);
    await command.action(logger, { options: { key: settingsNames.prompt, value: false } });
    assert.strictEqual(actualKey, settingsNames.prompt, 'Invalid key');
    assert.strictEqual(actualValue, false, 'Invalid value');
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
});
