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
const command: Command = require('./navigation-node-set');

describe(commands.NAVIGATION_NODE_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/team-a';
  const id = 2000;
  const nodeUrl = '/sites/team-a/sitepages/about.aspx';
  const title = 'About';
  const audienceIds = '7aa4a1ca-4035-4f2f-bac7-7beada59b5ba,4bbf236f-a131-4019-b4a2-315902fcfa3a';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
      request.patch
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
    assert.strictEqual(command.name.startsWith(commands.NAVIGATION_NODE_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly updates existing navigation node', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`
        && opts.data['Title'] === title
        && opts.data['Url'] === nodeUrl
        && JSON.stringify(opts.data['AudienceIds']) === JSON.stringify(audienceIds.split(','))
        && opts.data['IsExternal'] === false) {
        return;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { webUrl: webUrl, id: id, title: title, url: nodeUrl, isExternal: false, audienceIds: audienceIds } } as any);
    assert(patchStub.called);
  });

  it('correctly clears audienceIds from existing navigation node', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`
        && JSON.stringify(opts.data['AudienceIds']) === JSON.stringify([])) {
        return;
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { webUrl: webUrl, id: id, audienceIds: "" } } as any);
    assert(patchStub.called);
  });

  it('correctly handles API error', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/navigation/GetNodeById(${id})`) {
        throw { error: 'An error has occurred' };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, id: id, title: title, verbose: true } } as any), new CommandError('An error has occurred'));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'invalid', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is no options are set to be changed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if audienceIds contains more than 10 guids', async () => {
    const manyAudienceIds = `${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds},${audienceIds}`;
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, audienceIds: manyAudienceIds } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if audienceIds contains invalid guid', async () => {
    const invalidAudienceIds = `${audienceIds},invalid`;
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, audienceIds: invalidAudienceIds } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all options are set properly', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, title: title, url: nodeUrl, isExternal: true, audienceIds: audienceIds } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
