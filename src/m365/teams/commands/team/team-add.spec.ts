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
const command: Command = require('./team-add');

describe(commands.TEAM_ADD, () => {
  let log: string[];
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      global.setTimeout
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
    assert.strictEqual(command.name.startsWith(commands.TEAM_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation if name and description are passed when no template is passed', async () => {
    const actual = await command.validate({
      options: {
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if name and description are not passed when a template is supplied', async () => {
    const actual = await command.validate({
      options: {
        template: `abc`
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if description is not passed when no template is supplied', async () => {
    const actual = await command.validate({
      options: {
        name: 'Architecture'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if name is not passed when no template is supplied', async () => {
    const actual = await command.validate({
      options: {
        description: 'Architecture Discussion'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('creates Microsoft Teams team in the tenant when no template is supplied (verbose)', async () => {
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
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
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
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
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
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
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
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
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
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
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return Promise.resolve({
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
          });
        }
        return Promise.reject('Invalid request');
      });
    getRequestStub.onCall(1).callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
    });
    getRequestStub.onCall(2).callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
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

    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return Promise.resolve({
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
          });
        }
        return Promise.reject('Invalid request');
      });
    getRequestStub.onCall(1).callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
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
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get');
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return Promise.resolve({
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
          });
        }
        return Promise.reject('Invalid request');
      });
    getRequestStub.onCall(1).callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
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
});
