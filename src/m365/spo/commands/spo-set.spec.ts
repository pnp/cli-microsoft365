import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { Cli } from '../../../cli/Cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { CommandError } from '../../../Command.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command from './spo-set.js';
import { centralizedAfterEachHook, centralizedAfterHook, centralizedBeforeEachHook, centralizedBeforeHook, logger } from '../../../utils/tests.js';

describe(commands.SET, () => {
  let commandInfo: CommandInfo;

  before(() => {
    centralizedBeforeHook();
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    centralizedBeforeEachHook();
  });

  afterEach(() => {
    centralizedAfterEachHook();
    auth.service.spoUrl = undefined;
  });

  after(() => {
    centralizedAfterHook();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets SPO URL when no URL was set previously', async () => {
    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(auth.service.spoUrl, 'https://contoso.sharepoint.com');
  });

  it('sets SPO URL when other URL was set previously', async () => {
    auth.service.spoUrl = 'https://northwind.sharepoint.com';

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(auth.service.spoUrl, 'https://contoso.sharepoint.com');
  });

  it('throws error when trying to set SPO URL when not logged in to M365', async () => {
    auth.service.connected = false;

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com' } } as any), new CommandError('Log in to Microsoft 365 first'));
    assert.strictEqual(auth.service.spoUrl, undefined);
  });

  it('throws error when setting the password fails', async () => {
    auth.service.connected = true;
    sinonUtil.restore(auth.storeConnectionInfo);
    sinon.stub(auth, 'storeConnectionInfo').rejects(new Error('An error has occurred while setting the password'));

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com' } } as any), new CommandError('An error has occurred while setting the password'));
  });

  it('supports specifying url', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--url') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if url is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/team-a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

});
