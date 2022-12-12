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
const command: Command = require('./file-checkin');

describe(commands.FILE_CHECKIN, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const stubPostResponses: any = (getFileByServerRelativeUrlResp: any = null, getFileByIdResp: any = null) => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if (getFileByServerRelativeUrlResp) {
        return getFileByServerRelativeUrlResp;
      }
      else {
        if ((opts.url as string).indexOf('/_api/web/GetFileByServerRelativeUrl(') > -1) {
          return Promise.resolve();
        }
      }

      if (getFileByIdResp) {
        return getFileByIdResp;
      }
      else {
        if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });
  };

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
      request.post,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_CHECKIN), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('command correctly handles file get reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 'f09c4efe-b8c0-4e89-a166-03418661b89b'
      }
    }), new CommandError(err));
  });

  it('should handle checkin with url promise rejection', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File Not Found." } } });
    const getFileByServerRelativeUrlResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(getFileByServerRelativeUrlResp);

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }), new CommandError(expectedError));
  });

  it('should handle checkin with id promise rejection', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File Not Found." } } });
    const getFileByIdResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, getFileByIdResp);

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }), new CommandError(expectedError));
  });

  it('should call the correct API url when UniqueId option is passed', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    await command.action(logger, {
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/checkin(comment=\'\',checkintype=1)');
  });

  it('should call the correct API url when URL option is passed', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        url: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx')/checkin(comment='',checkintype=1)");
  });

  it('should call the correct API url when tenant root URL option is passed', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        url: '/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/_api/web/GetFileByServerRelativeUrl('%2FDocuments%2FTest1.docx')/checkin(comment='',checkintype=1)");
  });

  it('should call correctly the API when type is minor', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        url: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        type: 'minor'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx')/checkin(comment='',checkintype=0)");
  });

  it('should call correctly the API when type is overwrite', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        url: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        type: 'overwrite'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx')/checkin(comment='',checkintype=2)");
  });

  it('should call correctly the API when comment specified', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        url: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        comment: 'abc1'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativeUrl('%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx')/checkin(comment='abc1',checkintype=1)");
  });

  it('should call correctly the API when type is minor (id)', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        type: 'minor'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/checkin(comment='',checkintype=0)");
  });

  it('should call correctly the API when type is overwrite (id)', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        type: 'overwrite'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/checkin(comment='',checkintype=2)");
  });

  it('should call correctly the API when comment specified (id)', async () => {
    const postStub: sinon.SinonStub = stubPostResponses();

    await command.action(logger, {
      options: {
        id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        comment: 'abc1'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, "https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')/checkin(comment='abc1',checkintype=1)");
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

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id or url option not specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when comment lenght more than 1023', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', comment: 'ayMfuJBMDa3y3y8qitRb4U6VSbBVjeIxno45Ws6baZ1uatvxGVDS98zQu88QUjyeYXDbLey1dWTKdgMGw4LGeqfn080VszB5vMCrBEAnLYT54E94eW3YQe67Ub92oD0DG0U8gxMQJ0SWdVG9m5R5dL31YWx1Y5OH8KMtoAFkfo2lnbHVBMnCiO8oyuiRzVbTLkZB7mdih3F74ck3kEM7Lr1ayXkwHKK5h9MnTcVTWZVXafMOsuLYaVnUB7auhaamQ4JMBUFNpKhCjrNQVlYz0NlwJimlk5tPeR6crgeCm3u4YJtc1dBL2Ex7FRfvJ4g44WnkPLyU3PIXrHTjZtlgOKn4m9BiABuwznqiuytCcKbLxaTQcqHsbC7w20vnZxnLHYNnqXeDqwf6o43Si6duSeIZSixwoK4nE8qpCZk36jkwZBXASuv5aOyWLOsD19JjK7Ev3567b6oo11krIOpd0TSRihphELWnk9A71xpkCN1ljmSTnrITgQ7AxIaWOHvBIv5Swffi6AUM2DeLyz61EVe0fgAdVU3UySGSHGmUJEGqVBUlX7zZw2xSWswgvQphziHp2sKcnONWaaeDvbr27g67HrkkyYO3z0R5nY9TdSfkqDDQVSFdM3Sd6WLRKKKn64pcUzo9NcFNKzMSvRR0FbZFirpEcIfTCrSLaIiRZYCoGdj0BfePz83cimDmlVWS87UXugXmeWNpKTqQ1qG9y0fMwGIxFory4YbeRP9vKqX0vueCGKErb7tItC09jpLp0J8yMaj0iDdZ83Yc3JHunVmqZh56hmUroU8ER6ApPS3oDooEGH59e5I4DU8LG4rpAPmECX6oC8w9eZfM7U0uugT9Yx2ZAoDwvk0jYJz8SuU6dL6aFYtf7wzYcBcjc8gBySbeVZYPoLE3TGP1A0K8HNiZavHjsJWK0GIYVDT4QEsJO4R9PykRkn0O6TyDkaIgqju9hV7lqy9YqKawvBAUlNyK7b01fkra5UBrZzYz83k3OYWmG2naAcKuNuPs7OJ6' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation wrong type', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', type: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and url options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', url: '/sites/project-x/documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation id, type and comment params correct', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', type: 'overwrite', comment: '123' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation url, type and comment params correct', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', url: '/sites/docs/abc.txt', type: 'overwrite', comment: '123' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation type is major', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', url: '/sites/docs/abc.txt', type: 'major' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation type is minor', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', url: '/sites/docs/abc.txt', type: 'minor' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
