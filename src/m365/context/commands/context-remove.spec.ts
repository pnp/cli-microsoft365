import * as assert from 'assert';
import * as sinon from 'sinon';
import commands from '../commands';
import { Logger } from '../../../cli/Logger';
import { sinonUtil } from '../../../utils/sinonUtil';
import appInsights from '../../../appInsights';
import Command from '../../../Command';
import * as fs from 'fs';
const command: Command = require('./context-remove');

describe(commands.REMOVE, () => {
  let log: any[];
  let logger: Logger;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent,
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes a context from the .m365rc.json file', async () => {
    await command.action(logger, { options: { verbose: true } });

    let fileContent;
    let contextInfo;

    const fileExist = fs.existsSync('.m365rc.json');
    if (fileExist) {
      fileContent = fs.readFileSync('.m365rc.json', 'utf8');
      const contextInfoParsed = JSON.parse(fileContent);
      contextInfo = contextInfoParsed.context;
    }

    assert.deepStrictEqual(contextInfo, undefined);
  });

});