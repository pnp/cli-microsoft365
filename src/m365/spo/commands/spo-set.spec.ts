import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../telemetry';
import auth from '../../../Auth';
import { Cli } from '../../../cli/Cli';
import { CommandInfo } from '../../../cli/CommandInfo';
import { Logger } from '../../../cli/Logger';
import Command, { CommandError } from '../../../Command';
import { pid } from '../../../utils/pid';
import { session } from '../../../utils/session';
import { sinonUtil } from '../../../utils/sinonUtil';
import commands from '../commands';
const command: Command = require('./spo-set');

describe(commands.SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
  });

  afterEach(() => {
    auth.service.spoUrl = undefined;
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('sets SPO URL when no URL was set previously', async () => {
    auth.service.spoUrl = undefined;

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(auth.service.spoUrl, 'https://contoso.sharepoint.com');
  });

  it('sets SPO URL when other URL was set previously', async () => {
    auth.service.spoUrl = 'https://northwind.sharepoint.com';

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(auth.service.spoUrl, 'https://contoso.sharepoint.com');
  });

  it('throws error when trying to set SPO URL when not logged in to O365', async () => {
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
