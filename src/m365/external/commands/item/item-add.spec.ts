import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './item-add.js';

describe(commands.ITEM_ADD, () => {
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.put
    ]);
    loggerLogSpy.resetHistory();
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ITEM_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds an external item with simple properties', async () => {
    const externalItem = {
      "id": "ticket1",
      "acl": [
        {
          "type": "everyone",
          "value": "everyone",
          "accessType": "grant"
        }
      ],
      "properties": {
        "ticketTitle": "Something went wrong ticket",
        "priority": "high",
        "assignee": "Steve"
      },
      "content": {
        "value": "Something went wrong",
        "type": "text"
      },
      "activities": []
    };
    sinon.stub(request, 'put').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/connection/items/ticket1`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'ticket1',
      externalConnectionId: 'connection',
      content: 'Something went wrong',
      contentType: 'text',
      acls: 'grant,everyone,everyone',
      ticketTitle: 'Something went wrong ticket',
      priority: 'high',
      assignee: 'Steve'
    };
    await command.action(logger, { options } as any);
    assert(loggerLogSpy.calledWith(externalItem));
  });

  it('adds an external item with a collection properties', async () => {
    const externalItem = {
      "id": "ticket1",
      "acl": [
        {
          "type": "group",
          "value": "Admins",
          "accessType": "grant"
        }
      ],
      "properties": {
        "ticketTitle": "Something went wrong ticket",
        "priority": "high",
        "assignee@odata.type": "Collection(String)",
        "assignee": ["Steve", "Brian"]
      },
      "content": {
        "value": "Something went wrong",
        "type": "text"
      },
      "activities": []
    };
    sinon.stub(request, 'put').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/connection/items/ticket1`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'ticket1',
      externalConnectionId: 'connection',
      content: 'Something went wrong',
      acls: 'grant,group,Admins',
      ticketTitle: 'Something went wrong ticket',
      priority: 'high',
      'assignee@odata.type': 'Collection(String)',
      assignee: 'Steve;#Brian'
    };
    await command.action(logger, { options } as any);
    assert(loggerLogSpy.calledWith(externalItem));
  });

  it('outputs properties in csv output', async () => {
    const externalItem = {
      "id": "ticket1",
      "acl": [
        {
          "type": "everyone",
          "value": "everyone",
          "accessType": "grant"
        }
      ],
      "properties": {
        "ticketTitle": "Something went wrong ticket",
        "priority": "high",
        "assignee@odata.type": "Collection(String)",
        "assignee": [
          "Steve",
          "Brian"
        ]
      },
      "content": {
        "value": "Something went wrong",
        "type": "text"
      },
      "activities": []
    };
    sinon.stub(request, 'put').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/connection/items/ticket1`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'ticket1',
      externalConnectionId: 'connection',
      content: 'Something went wrong',
      contentType: 'text',
      acls: 'grant,everyone,everyone',
      ticketTitle: 'Something went wrong ticket',
      priority: 'high',
      "assignee@odata.type": "Collection(String)",
      assignee: 'Steve;#Brian',
      output: 'csv'
    };
    await command.action(logger, { options } as any);
    const extendedItem = {
      "id": "ticket1",
      "acl": [
        {
          "type": "everyone",
          "value": "everyone",
          "accessType": "grant"
        }
      ],
      "properties": {
        "ticketTitle": "Something went wrong ticket",
        "priority": "high",
        "assignee@odata.type": "Collection(String)",
        "assignee": ["Steve", "Brian"]
      },
      "content": {
        "value": "Something went wrong",
        "type": "text"
      },
      "activities": [],
      "ticketTitle": "Something went wrong ticket",
      "priority": "high",
      "assignee@odata.type": "Collection(String)",
      "assignee": "Steve, Brian"
    };
    assert(loggerLogSpy.calledWith(extendedItem));
  });

  it('outputs properties in md output', async () => {
    const externalItem = {
      "id": "ticket1",
      "acl": [
        {
          "type": "everyone",
          "value": "everyone",
          "accessType": "grant"
        }
      ],
      "properties": {
        "ticketTitle": "Something went wrong ticket",
        "priority": "high",
        "assignee": "Steve"
      },
      "content": {
        "value": "Something went wrong",
        "type": "text"
      },
      "activities": []
    };
    sinon.stub(request, 'put').callsFake(async (opts: any) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/external/connections/connection/items/ticket1`) {
        return externalItem;
      }
      throw 'Invalid request';
    });
    const options: any = {
      id: 'ticket1',
      externalConnectionId: 'connection',
      content: 'Something went wrong',
      contentType: 'text',
      acls: 'grant,everyone,everyone',
      ticketTitle: 'Something went wrong ticket',
      priority: 'high',
      assignee: 'Steve',
      output: 'md'
    };
    await command.action(logger, { options } as any);
    const extendedItem = {
      "id": "ticket1",
      "acl": [
        {
          "type": "everyone",
          "value": "everyone",
          "accessType": "grant"
        }
      ],
      "properties": {
        "ticketTitle": "Something went wrong ticket",
        "priority": "high",
        "assignee": "Steve"
      },
      "content": {
        "value": "Something went wrong",
        "type": "text"
      },
      "activities": [],
      "ticketTitle": "Something went wrong ticket",
      "priority": "high",
      "assignee": "Steve"
    };
    assert(loggerLogSpy.calledWith(extendedItem));
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'put').callsFake(() => {
      throw {
        "error": {
          "code": "Error",
          "message": "An error has occurred",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      };
    });

    const options: any = {
      id: 'ticket1',
      externalConnectionId: 'connection',
      content: 'Something went wrong',
      contentType: 'text',
      acls: 'grant,everyone,everyone',
      ticketTitle: 'Something went wrong ticket',
      priority: 'high',
      assignee: 'Steve'
    };
    await assert.rejects(command.action(logger, { options } as any),
      new CommandError(`An error has occurred`));
  });

  //#region validation
  it('fails validation when invalid contentType specified', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'invalid',
        acls: 'grant,everyone,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('passes validation when contentType is text', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,everyone,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when contentType is html', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'html',
        acls: 'grant,everyone,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when one acl with other than 3 elements', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('fails validation when multiple acls specified where one is with other than 3 elements', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,everyone,everyone;grant,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('passes validation for a single correct acl', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,everyone,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for multiple correct acls', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,everyone,everyone;grant,everyone,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation for invalid acl access type', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'invalid,everyone,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('passes validation for acl access type grant', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,everyone,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for acl access type deny', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'deny,everyone,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation for invalid acl type', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,invalid,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, false);
  });

  it('passes validation for acl type user', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,user,steve@contoso.com',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for acl type grant', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,group,Users',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for acl type everyone', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,everyone,everyone',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for acl type everyoneExceptGuests', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,everyoneExceptGuests,everyoneExceptGuests',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for acl type externalGroup', async () => {
    const actual = await command.validate({
      options: {
        id: 'ticket1',
        externalConnectionId: 'connection',
        content: 'Hello world',
        contentType: 'text',
        acls: 'grant,externalGroup,Users',
        name: 'Test item'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
  //#endregion

  it('allows unknown properties', () => {
    const allowUnknownOptions = command.allowUnknownOptions();
    assert.strictEqual(allowUnknownOptions, true);
  });

  //#region options
  it('supports specifying id', () => {
    const containsOption = command.options
      .some(o => o.option.indexOf('--id') > -1);
    assert(containsOption);
  });

  it('supports specifying externalConnectionId', () => {
    const containsOption = command.options
      .some(o => o.option.indexOf('--externalConnectionId') > -1);
    assert(containsOption);
  });

  it('supports specifying content', () => {
    const containsOption = command.options
      .some(o => o.option.indexOf('--content') > -1);
    assert(containsOption);
  });

  it('supports specifying contentType', () => {
    const containsOption = command.options
      .some(o => o.option.indexOf('--contentType') > -1);
    assert(containsOption);
  });

  it('supports specifying acls', () => {
    const containsOption = command.options
      .some(o => o.option.indexOf('--acls') > -1);
    assert(containsOption);
  });
  //#endregion
});
