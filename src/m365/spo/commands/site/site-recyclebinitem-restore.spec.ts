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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./site-recyclebinitem-restore');

describe(commands.SITE_RECYCLEBINITEM_RESTORE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
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
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_RECYCLEBINITEM_RESTORE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the siteUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo', ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ids option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '9526' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the second id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526, 9526' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the siteUrl and ids options are valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if siteUrl and id are defined', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when multiple IDs are specified', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526,5fb84a1f-6ab5-4d07-a6aa-31bba6de9527' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when multiple IDs with a space after the comma are specified', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526, 5fb84a1f-6ab5-4d07-a6aa-31bba6de9527' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('restores specified items from the recycle bin', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/site/RecycleBin/RestoreByIds') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const result = await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526,1adcf0d6-3733-4c13-b883-c84a27905cfd'
      }
    });

    assert.equal(result, undefined);
  });

  it('catches error when restores all items from recyclebin', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526,1adcf0d6-3733-4c13-b883-c84a27905cfd'
      }
    } as any), new CommandError('Invalid request'));
  });

  it('verifies that the command will fail when one of the promises fails', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.data.ids).filter((chunk: string) => chunk === 'fail').length > 0) {
        throw 'Invalid item';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9412, 1adcf0d6-3733-4c13-b883-c84a27905af4, fail, 641e5c65-a981-4910-b094-c212115b6d54, 5fb84a1f-6ab5-4d07-a6aa-31bba6de9526, 1adcf0d6-3733-4c13-b883-c84a27905cfd, 241e5c65-a981-4910-b094-c212115b6d5f, dc25898c-c977-4443-a821-5535e852975f, ccfb360c-7804-4e81-9cc8-8ea1a4fa53e0, a7598f93-7a7e-45c8-84db-7071bfec2840, 67786192-76b4-42f4-a8e3-aa0c5b00f96b, 5d32c945-a4a9-4b61-94ab-5de7095b2322, 241e5c65-a981-4910-b094-c212115b6d5f, dc25898c-c977-4443-a821-5535e852975f, ccfb360c-7804-4e81-9cc8-8ea1a4fa53e0, a7598f93-7a7e-45c8-84db-7071bfec2840, 67786192-76b4-42f4-a8e3-aa0c5b00f96b, 5d32c945-a4a9-4b61-94ab-5de7095b2322, 241e5c65-a981-4910-b094-c212115b6d5f, dc25898c-c977-4443-a821-5535e852975f, ccfb360c-7804-4e81-9cc8-8ea1a4fa53e0, a7598f93-7a7e-45c8-84db-7071bfec2840, 67786192-76b4-42f4-a8e3-aa0c5b00f96b'
      }
    }), new CommandError('Invalid item'));
  });

  it('restores specified items from the recycle bin in multiple chunks', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/site/RecycleBin/RestoreByIds') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const result = await command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9412, 1adcf0d6-3733-4c13-b883-c84a27905af4, 641e5c65-a981-4910-b094-c212115b6d54, 5fb84a1f-6ab5-4d07-a6aa-31bba6de9526, 1adcf0d6-3733-4c13-b883-c84a27905cfd, 241e5c65-a981-4910-b094-c212115b6d5f, dc25898c-c977-4443-a821-5535e852975f, ccfb360c-7804-4e81-9cc8-8ea1a4fa53e0, a7598f93-7a7e-45c8-84db-7071bfec2840, 67786192-76b4-42f4-a8e3-aa0c5b00f96b, 5d32c945-a4a9-4b61-94ab-5de7095b2322, 241e5c65-a981-4910-b094-c212115b6d5f, dc25898c-c977-4443-a821-5535e852975f, ccfb360c-7804-4e81-9cc8-8ea1a4fa53e0, a7598f93-7a7e-45c8-84db-7071bfec2840, 67786192-76b4-42f4-a8e3-aa0c5b00f96b, 5d32c945-a4a9-4b61-94ab-5de7095b2322, 241e5c65-a981-4910-b094-c212115b6d5f, dc25898c-c977-4443-a821-5535e852975f, ccfb360c-7804-4e81-9cc8-8ea1a4fa53e0, a7598f93-7a7e-45c8-84db-7071bfec2840, 67786192-76b4-42f4-a8e3-aa0c5b00f96b, 5d32c945-a4a9-4b61-94ab-5de7095b2322'
      }
    });

    assert.equal(result, undefined);
  });
});
