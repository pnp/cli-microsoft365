import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './team-add.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { entraUser } from '../../../../utils/entraUser.js';

describe(commands.TEAM_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        accessToken: 'abc'
      };
    }
    commandInfo = cli.getCommandInfo(command);
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
    (command as any).pollingInterval = 0;
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      entraGroup.getGroupById,
      entraUser.getUserIdByUpn,
      entraUser.getUserIdByEmail,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TEAM_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation if the template is not set and name and description is passed', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if ownerUserNames is set and it contains one or more invalid user principal names', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', ownerUserNames: 'invalid,john@contoso.com,doe@contoso.com,kevin@contoso' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ownerEmails is set and it contains one or more invalid user principal names', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', ownerEmails: 'invalid,john@contoso.com,doe@contoso.com,kevin@contoso' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ownerIds is set and it contains one or more invalid GUIDs', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', ownerIds: 'invalid,f5332379-663f-49b7-b5c6-84424ab9a0d1,80fcda19-6c95-4c58-bc98-bc9dfb49bd0d' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if memberUserNames is set and it contains one or more invalid user principal names', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', memberUserNames: 'invalid,john@contoso.com,doe@contoso.com,kevin@contoso' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if memberEmails is set and it contains one or more invalid user principal names', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', memberEmails: 'invalid,john@contoso.com,doe@contoso.com,kevin@contoso' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if memberIds is set and it contains one or more invalid GUIDs', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', memberIds: 'invalid,f5332379-663f-49b7-b5c6-84424ab9a0d1,80fcda19-6c95-4c58-bc98-bc9dfb49bd0d' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if ownerUserNames is set and it contains valid userNames', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', ownerUserNames: 'john@contoso.com,doe@contoso.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if ownerEmails is set and it contains valid userNames', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', ownerEmails: 'john@contoso.com,doe@contoso.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if ownerIds is set and it contains valid GUIDs', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', ownerIds: 'f5332379-663f-49b7-b5c6-84424ab9a0d1,80fcda19-6c95-4c58-bc98-bc9dfb49bd0d' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if memberUserNames is set and it contains valid userNames', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', memberUserNames: 'john@contoso.com,doe@contoso.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if memberEmails is set and it contains valid userNames', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', memberEmails: 'john@contoso.com,doe@contoso.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if memberIds is set and it contains valid GUIDs', async () => {
    const actual = await command.validate({ options: { name: 'TeamName', description: 'Description', memberIds: 'f5332379-663f-49b7-b5c6-84424ab9a0d1,80fcda19-6c95-4c58-bc98-bc9dfb49bd0d' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws error when trying to create a team using application permissions and not specifying an owner', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {
        name: 'TeamName',
        description: 'Description'
      }
    }), new CommandError(`Specify at least 'ownerUserNames', 'ownerIds' or 'ownerEmails' when using application permissions.`));
  });

  it('creates Microsoft Teams team in the tenant when template is supplied and will continue fetching aadGroup when error is being thrown when wait is set to true', async () => {
    const groupId = '79afc64f-c76b-4edc-87f3-a47a1264695a';
    let firstCall = true;

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/teams') {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return { "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity", "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1", "operationType": "createTeam", "createdDateTime": "2020-06-15T22:28:16.3007846Z", "status": "succeeded", "lastActionDateTime": "2020-06-15T22:28:16.3007846Z", "attemptsCount": 1, "targetResourceId": groupId, "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')", "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"Succeeded\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}", "error": null };
      }
      throw 'Invalid request';
    });
    const aadGroupStub = sinon.stub(entraGroup, 'getGroupById').callsFake(async (groupId: string) => {
      if (firstCall) {
        firstCall = false;
        throw {
          code: 'Request_ResourceNotFound',
          message: "Resource '4deeae0d-0402-4a08-b2cf-fe9f060fb625' does not exist or one of its queried reference-property objects are not present.",
          innerError: [Object]
        };
      }
      else {
        return { "id": groupId, "deletedDateTime": null, "classification": null, "createdDateTime": "2023-05-23T20:15:43Z", "creationOptions": ["Team", "ExchangeProvisioningFlags:3552"], "description": "TEST", "displayName": "CLITEST2", "expirationDateTime": null, "groupTypes": ["Unified"], "isAssignableToRole": null, "mail": "CLITEST2691@mathijsdev2.onmicrosoft.com", "mailEnabled": true, "mailNickname": "CLITEST2691", "membershipRule": null, "membershipRuleProcessingState": null, "onPremisesDomainName": null, "onPremisesLastSyncDateTime": null, "onPremisesNetBiosName": null, "onPremisesSamAccountName": null, "onPremisesSecurityIdentifier": null, "onPremisesSyncEnabled": null, "preferredDataLocation": null, "preferredLanguage": null, "proxyAddresses": ["SMTP:CLITEST2691@mathijsdev2.onmicrosoft.com"], "renewedDateTime": "2023-05-23T20:15:43Z", "resourceBehaviorOptions": ["HideGroupInOutlook", "SubscribeMembersToCalendarEventsDisabled", "WelcomeEmailDisabled"], "resourceProvisioningOptions": ["Team"], "securityEnabled": false, "securityIdentifier": "S-1-12-1-4197740414-1080776454-10446780-3069961820", "theme": null, "visibility": "Public", "onPremisesProvisioningErrors": [] };
      }
    });

    await command.action(logger, {
      options: {
        verbose: true,
        template: '{"template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates(\'standard\')"}',
        wait: true
      }
    });
    assert(aadGroupStub.calledTwice);
  });

  it('creates Microsoft Teams team in the tenant when no template is supplied (verbose)', async () => {
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "inProgress",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    });
    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Architecture',
      description: 'Architecture Discussion'
    });
    assert(getRequestStub.called);
  });

  it('creates Microsoft Teams team in the tenant when template is supplied (verbose)', async () => {
    const template = `
    {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`;
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "inProgress",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        template
      }
    });
    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Sample Engineering Team',
      description: 'This is a sample engineering team, used to showcase the range of properties supported by this API'
    });
    assert(getRequestStub.called);
  });

  it('creates Microsoft Teams team in the tenant when template and name is supplied (verbose)', async () => {
    const template = `
    {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`;
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "inProgress",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        name: 'Sample Classroom Team',
        template
      }
    });
    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Sample Classroom Team',
      description: 'This is a sample engineering team, used to showcase the range of properties supported by this API'
    });
    assert(getRequestStub.called);
  });

  it('creates Microsoft Teams team in the tenant when template and description is supplied (verbose)', async () => {
    const template = `
    {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`;
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "inProgress",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        template
      }
    });
    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Sample Engineering Team',
      description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
    });
    assert(getRequestStub.called);
  });

  it('creates Microsoft Teams team in the tenant when template, name and description is supplied (verbose)', async () => {
    const template = `
    {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`;
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "inProgress",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        template
      }
    });
    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Sample Classroom Team',
      description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
    });
    assert(getRequestStub.called);
  });

  it('creates Microsoft Teams team in the tenant when template, name and description is supplied and waits for command to complete (verbose)', async () => {
    const template = `
    {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`;
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
            "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
            "operationType": "createTeam",
            "createdDateTime": "2020-06-15T22:28:16.3007846Z",
            "status": "inProgress",
            "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
            "attemptsCount": 1,
            "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
            "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
            "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"Failed\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
            "error": null
          };
        }
        throw 'Invalid request';
      });
    getRequestStub.onCall(1).callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "inProgress",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"Failed\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });
    getRequestStub.onCall(2).callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "succeeded",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"Failed\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });



    await command.action(logger, {
      options: {
        verbose: true,
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        template
      }
    });
    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Sample Classroom Team',
      description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
    });
  });

  it('correctly handles error when creating a Team', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles operation error when creating a Team when waiting for command to complete', async () => {
    const template = `
    {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`;

    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
            "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
            "operationType": "createTeam",
            "createdDateTime": "2020-06-15T22:28:16.3007846Z",
            "status": "inProgress",
            "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
            "attemptsCount": 1,
            "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
            "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
            "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"Failed\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
            "error": null
          };
        }
        throw 'Invalid request';
      });
    getRequestStub.onCall(1).callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "failed",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"Failed\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": 'An error has occurred'
        };
      }
      throw 'Invalid request';
    });


    await assert.rejects(command.action(logger, {
      options: {
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        template
      }
    } as any), new CommandError('An error has occurred'));

    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Sample Classroom Team',
      description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
    });
  });

  it('correctly handles inProgress operation status when creating a Team and waiting for the command to complete', async () => {
    const template = `
    {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`;
    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake(async (opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
            "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
            "operationType": "createTeam",
            "createdDateTime": "2020-06-15T22:28:16.3007846Z",
            "status": "inProgress",
            "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
            "attemptsCount": 1,
            "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
            "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
            "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"Failed\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
            "error": null
          };
        }
        throw 'Invalid request';
      });
    getRequestStub.onCall(1).callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "succeeded",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"Failed\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        template
      }
    } as any);
    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Sample Classroom Team',
      description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
    });
    assert(loggerLogSpy.called);
  });

  it('creates Microsoft Teams team in the tenant when using application only permissions and specifying owners and members by email', async () => {
    const userId = 'd0fe0abd-bfe8-4a7d-9957-e8fb2d739f61';
    const userId2 = 'bed5c7fa-25cd-47aa-9006-566d2c4813ec';
    const groupId = '8d91e4d3-b02f-4c26-b1bc-f15f5241f539';

    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${groupId}/members` && opts.data['user@odata.bind'].indexOf(userId2) > -1) {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "succeeded",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });

    sinon.stub(entraGroup, 'getGroupById').resolves({ 'id': groupId, displayName: 'Architecture' });
    sinon.stub(entraUser, 'getUserIdsByEmails').resolves([userId, userId2]);
    sinon.stub(entraUser, 'getUserIdsByUpns').resolves([userId]);

    await command.action(logger, {
      options: {
        verbose: true,
        name: 'Architecture',
        description: 'Architecture Discussion',
        ownerEmails: 'john@contoso.com,doe@contoso.com',
        memberUserNames: 'john@contoso.com'
      }
    });
    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Architecture',
      description: 'Architecture Discussion',
      members: [{
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner', 'member'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`
      }]
    });
  });

  it('creates Microsoft Teams team in the tenant when using application only permissions and specifying owner by id', async () => {
    const userId = 'd0fe0abd-bfe8-4a7d-9957-e8fb2d739f61';

    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const requestStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return { statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
          "id": "8ad1effa-7ed1-4d03-bd60-fe177d8d56f1",
          "operationType": "createTeam",
          "createdDateTime": "2020-06-15T22:28:16.3007846Z",
          "status": "succeeded",
          "lastActionDateTime": "2020-06-15T22:28:16.3007846Z",
          "attemptsCount": 1,
          "targetResourceId": "79afc64f-c76b-4edc-87f3-a47a1264695a",
          "targetResourceLocation": "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')",
          "Value": "{\"apps\":[{\"Index\":1,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"com.microsoft.teamspace.tab.vsts\"},{\"Index\":2,\"Status\":\"InProgress\",\"UpdateTimestamp\":\"2020-06-15T22:28:16.8753199+00:00\",\"Reference\":\"1542629c-01b3-4a6d-8f76-1938b779e48d\"}],\"channels\":[{\"tabs\":[],\"Index\":1,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Class Announcements\"},{\"tabs\":[],\"Index\":2,\"Status\":\"NotStarted\",\"UpdateTimestamp\":\"2020-06-15T22:28:14.0279825+00:00\",\"Reference\":\"Homework\"}],\"WorkflowId\":\"northeurope.695866c1-c68a-435c-b707-432984ec721c\"}",
          "error": null
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true,
        name: 'Architecture',
        description: 'Architecture Discussion',
        ownerIds: userId
      }
    });
    assert.deepEqual(requestStub.getCall(0).args[0].data, {
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Architecture',
      description: 'Architecture Discussion',
      members: [{
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`
      }]
    });
  });
});