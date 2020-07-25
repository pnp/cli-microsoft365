import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./team-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';
import * as chalk from 'chalk';

describe(commands.TEAMS_TEAM_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      request.get,
      fs.existsSync,
      fs.readFileSync,
      global.setTimeout
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TEAMS_TEAM_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation if name and description are passed when no template is passed', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation if name and description are not passed when a template is supplied', (done) => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = (command.validate() as CommandValidate)({
      options: {
        templatePath: 'template.json'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('fails validation if description is not passed when no template is supplied', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'Architecture'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if name is not passed when no template is supplied', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        description: 'Architecture Discussion'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if template not found', (done) => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = (command.validate() as CommandValidate)({
      options: {
        templatePath: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('creates Microsoft Teams team in the tenant when no template is supplied (verbose)', (done) => {
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args[0].body, {
          "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
          displayName: 'Architecture',
          description: 'Architecture Discussion'
        });
        assert(getRequestStub.called);
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template is supplied (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        templatePath: 'template.json'
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args[0].body, {
          "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
          displayName: 'Sample Engineering Team',
          description: 'This is a sample engineering team, used to showcase the range of properties supported by this API'
        });
        assert(getRequestStub.called);
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template and name is supplied (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        name: 'Sample Classroom Team',
        templatePath: 'template.json'
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args[0].body, {
          "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
          displayName: 'Sample Classroom Team',
          description: 'This is a sample engineering team, used to showcase the range of properties supported by this API'
        });
        assert(getRequestStub.called);
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template and description is supplied (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json'
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args[0].body, {
          "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
          displayName: 'Sample Engineering Team',
          description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
        });
        assert(getRequestStub.called);
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template, name and description is supplied (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json'
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args[0].body, {
          "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
          displayName: 'Sample Classroom Team',
          description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
        });
        assert(getRequestStub.called);
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant when template, name and description is supplied and waits for command to complete (verbose)', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get')
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json',
      }
    }, () => {
      try {
        assert.deepEqual(requestStub.getCall(0).args[0].body, {
          "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
          displayName: 'Sample Classroom Team',
          description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
        });
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when creating a Team', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        verbose: true,
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles operation error when creating a Team when waiting for command to complete', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get')
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json',
      }
    }, (err?: any) => {
      try {
        assert.deepEqual(requestStub.getCall(0).args[0].body, {
          "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
          displayName: 'Sample Classroom Team',
          description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
        });
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles inProgress operation status when creating a Team and waiting for the command to complete', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => `
    {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "displayName": "Sample Engineering Team",
      "description": "This is a sample engineering team, used to showcase the range of properties supported by this API"
    }`);
    const requestStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({ statusCode: 202, headers: { location: "/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')" } });
      }
      return Promise.reject('Invalid request');
    });

    const getRequestStub = sinon.stub(request, 'get')
    getRequestStub.onCall(0)
      .callsFake((opts) => {
        if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
          return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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
      if (opts.url === `https://graph.microsoft.com/beta/teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations('8ad1effa-7ed1-4d03-bd60-fe177d8d56f1')`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('79afc64f-c76b-4edc-87f3-a47a1264695a')/operations/$entity",
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

    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        wait: true,
        name: 'Sample Classroom Team',
        description: 'This is a sample classroom team, used to showcase the range of properties supported by this API',
        templatePath: 'template.json',
      }
    }, (err?: any) => {
      try {
        assert.deepEqual(requestStub.getCall(0).args[0].body, {
          "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
          displayName: 'Sample Classroom Team',
          description: 'This is a sample classroom team, used to showcase the range of properties supported by this API'
        });
        assert(cmdInstanceLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});