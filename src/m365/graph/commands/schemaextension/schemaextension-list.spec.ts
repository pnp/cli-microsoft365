import commands from '../../commands';
import Command, { CommandOption, CommandValidate} from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./schemaextension-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SCHEMAEXTENSION_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.SCHEMAEXTENSION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });
  it('lists schema extensions', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`)> -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgBAAAAAAAAAA%3d%3d%23RT%3a1%23TRC%3a1%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKACLAQgAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
          "value": [
              {
                  "id": "adatumisv_exo2",
                  "description": "sample desccription",
                  "targetTypes": [
                      "Message"
                  ],
                  "status": "Available",
                  "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
                  "properties": [
                      {
                          "name": "p1",
                          "type": "String"
                      },
                      {
                          "name": "p2",
                          "type": "String"
                      }
                  ]
              }
          ]
      });
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{
                  "id": "adatumisv_exo2",
                  "description": "sample desccription",
                  "targetTypes": [
                      "Message"
                  ],
                  "status": "Available",
                  "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
                  "properties": [
                      {
                          "name": "p1",
                          "type": "String"
                      },
                      {
                          "name": "p2",
                          "type": "String"
                      }
                  ]
              }]
         ));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });
  it('lists two schema extensions', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`)> -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgBAAAAAAAAAA%3d%3d%23RT%3a1%23TRC%3a1%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKACLAQgAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
          "value": [
              {
                  "id": "adatumisv_exo2",
                  "description": "sample desccription",
                  "targetTypes": [
                      "Message"
                  ],
                  "status": "Available",
                  "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
                  "properties": [
                      {
                          "name": "p1",
                          "type": "String"
                      },
                      {
                          "name": "p2",
                          "type": "String"
                      }
                  ]
              },
              {
                "id": "adatumisv_exo3",
                "description": "sample desccription",
                "targetTypes": [
                    "Message"
                ],
                "status": "Available",
                "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
                "properties": [
                    {
                        "name": "p1",
                        "type": "String"
                    },
                    {
                        "name": "p2",
                        "type": "String"
                    }
                ]
            }
          ]
      });
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.args[0][1].id === 'adatumisv_exo3');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });
  it('lists schema extensions with filter options', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`$filter`)> -1) {
        return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgGAAAAAAAAAA%3d%3d%23RT%3a2%23TRC%3a2%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKAAaAIcAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
            "value": [
                {
                    "id": "adatumisv_courses",
                    "description": "Extension description",
                    "targetTypes": [
                        "User",
                        "Group"
                    ],
                    "status": "Available",
                    "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
                    "properties": [
                        {
                            "name": "id",
                            "type": "Integer"
                        },
                        {
                            "name": "name",
                            "type": "String"
                        },
                        {
                            "name": "type",
                            "type": "String"
                        }
                    ]
                }
            ]
        });
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        owner:'07d21ad2-c8f9-4316-a14a-347db702bd3c'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
              {
                  "id": "adatumisv_courses",
                  "description": "Extension description",
                  "targetTypes": [
                      "User",
                      "Group"
                  ],
                  "status": "Available",
                  "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
                  "properties": [
                      {
                          "name": "id",
                          "type": "Integer"
                      },
                      {
                          "name": "name",
                          "type": "String"
                      },
                      {
                          "name": "type",
                          "type": "String"
                      }
                  ]
              }
          ]
      ));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });
  it('lists schema extensions on the second page no page size given', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`$top`)> -1) {
        return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgGAAAAAAAAAA%3d%3d%23RT%3a2%23TRC%3a2%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKAAaAIcAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
            "value": [
                {
                    "id": "adatumisv_courses",
                    "description": "Extension description",
                    "targetTypes": [
                        "User",
                        "Group"
                    ],
                    "status": "Available",
                    "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
                    "properties": [
                        {
                            "name": "id",
                            "type": "Integer"
                        },
                        {
                            "name": "name",
                            "type": "String"
                        },
                        {
                            "name": "type",
                            "type": "String"
                        }
                    ]
                }
            ]
        });
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        pageNumber:1
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
              {
                  "id": "adatumisv_courses",
                  "description": "Extension description",
                  "targetTypes": [
                      "User",
                      "Group"
                  ],
                  "status": "Available",
                  "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
                  "properties": [
                      {
                          "name": "id",
                          "type": "Integer"
                      },
                      {
                          "name": "name",
                          "type": "String"
                      },
                      {
                          "name": "type",
                          "type": "String"
                      }
                  ]
              }
          ]
      ));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });
  it('lists schema extensions on the page size 1 second page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`$top`)> -1) {
        return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgGAAAAAAAAAA%3d%3d%23RT%3a2%23TRC%3a2%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKAAaAIcAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
            "value": [
                {
                    "id": "adatumisv_courses",
                    "description": "Extension description",
                    "targetTypes": [
                        "User",
                        "Group"
                    ],
                    "status": "Available",
                    "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
                    "properties": [
                        {
                            "name": "id",
                            "type": "Integer"
                        },
                        {
                            "name": "name",
                            "type": "String"
                        },
                        {
                            "name": "type",
                            "type": "String"
                        }
                    ]
                }
            ]
        });
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        pageNumber:1,
        pageSize:1
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
              {
                  "id": "adatumisv_courses",
                  "description": "Extension description",
                  "targetTypes": [
                      "User",
                      "Group"
                  ],
                  "status": "Available",
                  "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
                  "properties": [
                      {
                          "name": "id",
                          "type": "Integer"
                      },
                      {
                          "name": "name",
                          "type": "String"
                      },
                      {
                          "name": "type",
                          "type": "String"
                      }
                  ]
              }
          ]
      ));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });
  it('lists schema extensions(debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`)> -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgBAAAAAAAAAA%3d%3d%23RT%3a1%23TRC%3a1%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKACLAQgAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
          "value": [
              {
                  "id": "adatumisv_exo2",
                  "description": "sample desccription",
                  "targetTypes": [
                      "Message"
                  ],
                  "status": "Available",
                  "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
                  "properties": [
                      {
                          "name": "p1",
                          "type": "String"
                      },
                      {
                          "name": "p2",
                          "type": "String"
                      }
                  ]
              }
          ]
      });
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
              {
                  "id": "adatumisv_exo2",
                  "description": "sample desccription",
                  "targetTypes": [
                      "Message"
                  ],
                  "status": "Available",
                  "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
                  "properties": [
                      {
                          "name": "p1",
                          "type": "String"
                      },
                      {
                          "name": "p2",
                          "type": "String"
                      }
                  ]
              }
          ]
      ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('passes validation if the owner is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { owner: '68be84bf-a585-4776-80b3-30aa5207aa22' } });
    assert.strictEqual(actual, true);
  });
  it('fails validation if the owner is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { owner: '123' } });
    assert.notStrictEqual(actual, true);
  });
  it('fails validation if the status is not a valid status', () => {
    const actual = (command.validate() as CommandValidate)({ options: { status: 'test' } });
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the status is a valid status', () => {
    const actual = (command.validate() as CommandValidate)({ options: { status: 'InDevelopment' } });
    assert.strictEqual(actual, true);
  });
  it('fails validation if the pageNumber is not positive number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { pageNumber: '-1' } });
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the pageNumber is a positive number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { pageNumber: '2' } });
    assert.strictEqual(actual, true);
  });
  it('fails validation if the pageSize is not positive number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { pageSize: '-1' } });
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the pageSize is a positive number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { pageSize: '2' } });
    assert.strictEqual(actual, true);
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
