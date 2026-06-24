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
import command, { options } from './config-set.js';

describe(commands.CONFIG_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
    sinon.stub(telemetry, 'trackEvent').resolves();
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


  it('fails validation if specified key is invalid ', () => {
    const actual = commandOptionsSchema.safeParse({ key: 'invalid', value: 'false' });
    assert.strictEqual(actual.success, false);
  });

  it(`passes validation if setting is set to ${settingsNames.showHelpOnFailure} and value to true`, () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.showHelpOnFailure, value: 'true' });
    assert.strictEqual(actual.success, true);
  });

  it(`passes validation if setting is set to ${settingsNames.showHelpOnFailure} and value to false`, () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.showHelpOnFailure, value: 'false' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if specified output type is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.output, value: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation for output type text', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.output, value: 'text' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for output type json', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.output, value: 'json' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for output type csv', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.output, value: 'csv' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if specified authType is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.authType, value: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation for authType type deviceCode', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.authType, value: 'deviceCode' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for authType type browser', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.authType, value: 'browser' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for authType type certificate', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.authType, value: 'certificate' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for authType type password', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.authType, value: 'password' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for authType type identity', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.authType, value: 'identity' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for authType type secret', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.authType, value: 'secret' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if specified error output type is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.errorOutput, value: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation for error output stdout', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.errorOutput, value: 'stdout' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for error output stderr', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.errorOutput, value: 'stderr' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if specified help mode is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.helpMode, value: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation for help mode options', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.helpMode, value: 'options' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for help mode examples', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.helpMode, value: 'examples' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for help mode remarks', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.helpMode, value: 'remarks' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for help mode response', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.helpMode, value: 'response' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for help mode full', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.helpMode, value: 'full' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if specified promptListPageSize value is a string', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.promptListPageSize, value: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if specified promptListPageSize value is 0', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.promptListPageSize, value: '0' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if specified promptListPageSize value is negative', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.promptListPageSize, value: '-1' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation for number value in promptListPageSize', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.promptListPageSize, value: '10' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if specified help target is invalid', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.helpTarget, value: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation for help target web', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.helpTarget, value: 'web' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation for help target console', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.helpTarget, value: 'console' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if specified clientId is not a GUID', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.clientId, value: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if specified clientId is a GUID', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.clientId, value: '00000000-0000-0000-c000-000000000001' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if specified tenantId is not a GUID or common', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.tenantId, value: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if specified tenantId is a GUID', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.tenantId, value: '00000000-0000-0000-c000-000000000001' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if specified tenantId is common', () => {
    const actual = commandOptionsSchema.safeParse({ key: settingsNames.tenantId, value: 'common' });
    assert.strictEqual(actual.success, true);
  });
});
