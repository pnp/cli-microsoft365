import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./list-contenttype-default-set');

describe(commands.LIST_CONTENTTYPE_DEFAULT_SET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_CONTENTTYPE_DEFAULT_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('configures specified visible content type as default. List specified using Title. UniqueContentTypeOrder null', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder`) {
        return Promise.resolve('');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
        return Promise.resolve({
          "ContentTypeOrder": [
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            },
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            }
          ],
          "UniqueContentTypeOrder": null
        });
      }

      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/ContentTypes?$select=Id`) {
        return Promise.resolve({
          value: [
            {
              Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
            },
            {
              Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    });
  });

  it('configures specified visible content type as default. List specified using Title. UniqueContentTypeOrder null. Debug', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder`) {
        return Promise.resolve('');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
        return Promise.resolve({
          "ContentTypeOrder": [
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            },
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            }
          ],
          "UniqueContentTypeOrder": null
        });
      }

      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/ContentTypes?$select=Id`) {
        return Promise.resolve({
          value: [
            {
              Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
            },
            {
              Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('configures specified visible content type as default. List specified using ID. UniqueContentTypeOrder not null', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/RootFolder`) {
        return Promise.resolve('');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
        return Promise.resolve({
          "ContentTypeOrder": [
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            },
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            }
          ],
          "UniqueContentTypeOrder": [
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            },
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            }
          ]
        });
      }

      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/ContentTypes?$select=Id`) {
        return Promise.resolve({
          value: [
            {
              Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
            },
            {
              Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    });
  });

  it('configures specified visible content type as default. List specified using ID. UniqueContentTypeOrder not null. Debug', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/RootFolder`) {
        return Promise.resolve('');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
        return Promise.resolve({
          "ContentTypeOrder": [
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            },
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            }
          ],
          "UniqueContentTypeOrder": [
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            },
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            }
          ]
        });
      }

      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/ContentTypes?$select=Id`) {
        return Promise.resolve({
          value: [
            {
              Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
            },
            {
              Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it('configures specified invisible content type as default. List specified using Title. UniqueContentTypeOrder null', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder` &&
        opts.headers &&
        opts.headers['x-http-method'] === 'MERGE' &&
        JSON.stringify(opts.data) === JSON.stringify({
          UniqueContentTypeOrder: [
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            },
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            }
          ]
        })) {
        return Promise.resolve('');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
        return Promise.resolve({
          "ContentTypeOrder": [
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            }
          ],
          "UniqueContentTypeOrder": null
        });
      }

      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/ContentTypes?$select=Id`) {
        return Promise.resolve({
          value: [
            {
              Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
            },
            {
              Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    });
  });

  it('configures specified invisible content type as default. List specified using Title. UniqueContentTypeOrder null. Debug', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder` &&
        opts.headers &&
        opts.headers['x-http-method'] === 'MERGE' &&
        JSON.stringify(opts.data) === JSON.stringify({
          UniqueContentTypeOrder: [
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            },
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            }
          ]
        })) {
        return Promise.resolve('');
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
        return Promise.resolve({
          "ContentTypeOrder": [
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            }
          ],
          "UniqueContentTypeOrder": null
        });
      }

      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/ContentTypes?$select=Id`) {
        return Promise.resolve({
          value: [
            {
              Id: { "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E" }
            },
            {
              Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it(`doesn't configure content type as default if it's already set as default`, async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
        return Promise.resolve({
          "ContentTypeOrder": [
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            },
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            }
          ],
          "UniqueContentTypeOrder": null
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    });
  });

  it(`doesn't configure content type as default if it's already set as default. Debug`, async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
        return Promise.resolve({
          "ContentTypeOrder": [
            {
              "StringValue": "0x0104001A75DCE30BAC754AA5134C183CF7A92E"
            },
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            }
          ],
          "UniqueContentTypeOrder": null
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    });
    assert(loggerLogToStderrSpy.called);
  });

  it(`fails, if the specified web doesn't exist`, async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('Request failed with status code 404');
    });

    await assert.rejects(command.action(logger, {
      options: {
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    }), new CommandError('Request failed with status code 404'));
  });

  it(`fails, if the list specified by title doesn't exist`, async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('Request failed with status code 404');
    });

    await assert.rejects(command.action(logger, {
      options: {
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    }), new CommandError('Request failed with status code 404'));
  });

  it(`fails, if the specified content type not found in the list`, async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`) {
        return Promise.resolve({
          "ContentTypeOrder": [
            {
              "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550"
            }
          ],
          "UniqueContentTypeOrder": null
        });
      }

      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/GetByTitle('My%20List')/ContentTypes?$select=Id`) {
        return Promise.resolve({
          value: [
            {
              Id: { "StringValue": "0x01009C993C306A41A9419C8F5267B74D414F00FD8183595A9B79489F81D6075ADFB550" }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        listTitle: 'My List',
        webUrl: 'https://contoso.sharepoint.com',
        contentTypeId: '0x0104001A75DCE30BAC754AA5134C183CF7A92E'
      }
    }), new CommandError('Content type 0x0104001A75DCE30BAC754AA5134C183CF7A92E missing in the list. Add the content type to the list first and try again.'));
  });

  it('fails validation if neither listId nor listTitle are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', contentTypeId: '0x0120' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '12345', contentTypeId: '0x0120' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', contentTypeId: '0x0120' } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the listTitle option is passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Documents', contentTypeId: '0x0120' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both listId and listTitle options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', listTitle: 'Documents', contentTypeId: '0x0120' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('configures contentTypeId as string option', () => {
    const types = command.types;
    ['c', 'contentTypeId'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });
});