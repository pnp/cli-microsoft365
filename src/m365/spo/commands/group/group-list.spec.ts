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
const command: Command = require('./group-list');

describe(commands.GROUP_LIST, () => {
  const groupsResponse = [{
    Id: 15,
    Title: "Contoso Members",
    LoginName: "Contoso Members",
    "Description": "SharePoint Contoso",
    IsHiddenInUI: false,
    PrincipalType: 8
  }];
  const groupsResponseValue = {
    value: groupsResponse
  };
  const associatedGroupsResponse = {
    "AssociatedMemberGroup":
    {
      "Id": 6,
      "Title": "Site Members",
      "LoginName": "Site Members",
      "Description": "",
      "IsHiddenInUI": false,
      "PrincipalType": 8
    },
    "AssociatedOwnerGroup": {
      "Id": 7,
      "Title": "Site Owners",
      "LoginName": "Site Owners",
      "Description": "",
      "IsHiddenInUI": false,
      "PrincipalType": 8
    },
    "AssociatedVisitorGroup": {
      "Id": 8,
      "Title": "Site Visitors",
      "LoginName": "Site Visitors",
      "Description": "",
      "IsHiddenInUI": false,
      "PrincipalType": 8
    }
  };
  const associatedGroupsResponseText = [{
    "Id": 6,
    "Title": "Site Members",
    "LoginName": "Site Members",
    "Description": "",
    "IsHiddenInUI": false,
    "PrincipalType": 8,
    "Type": "AssociatedMemberGroup"
  },
  {
    "Id": 7,
    "Title": "Site Owners",
    "LoginName": "Site Owners",
    "Description": "",
    "IsHiddenInUI": false,
    "PrincipalType": 8,
    "Type": "AssociatedOwnerGroup"
  },
  {
    "Id": 8,
    "Title": "Site Visitors",
    "LoginName": "Site Visitors",
    "Description": "",
    "IsHiddenInUI": false,
    "PrincipalType": 8,
    "Type": "AssociatedVisitorGroup"
  }
  ];



  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.GROUP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'Title', 'LoginName', 'IsHiddenInUI', 'PrincipalType', 'Type']);
  });

  it('retrieves all site groups', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups') > -1) {
        return groupsResponseValue;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith(groupsResponse));
  });

  it('retrieves associated groups from the site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web?$expand=') > -1) {
        return JSON.stringify(associatedGroupsResponse);
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com',
        associatedGroupsOnly: true
      }
    });
    assert(loggerLogSpy.calledWith(JSON.stringify(associatedGroupsResponse)));
  });

  it('retrieves associated groups from the site with return type json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web?$expand=') > -1) {
        return JSON.stringify(associatedGroupsResponse);
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com',
        associatedGroupsOnly: true,
        output: 'json'
      }
    });
    assert(loggerLogSpy.calledWith(JSON.stringify(associatedGroupsResponse)));
  });

  it('retrieves associated groups from the site with return type text', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web?$expand=') > -1) {
        return associatedGroupsResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com',
        associatedGroupsOnly: true,
        output: 'text'
      }
    });
    assert(loggerLogSpy.calledWith(associatedGroupsResponseText));
  });

  it('command correctly handles group list reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/sitegroups') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(err));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });


  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation url is valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
