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
const command: Command = require('./notebook-add');

describe(commands.NOTEBOOK_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.NOTEBOOK_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both userId and userName options are passed', async () => {
    const actual = await command.validate({
      options:
      {
        name: 'Private Notebook',
        userId: '2609af39-7775-4f94-a3dc-0dd67657e900',
        userName: 'Documents'
      }
    }, commandInfo);
    assert.strictEqual(actual, 'Specify either userId or userName, but not both');
  });

  it('fails validation if both groupId and groupName options are passed', async () => {
    const actual = await command.validate({
      options:
      {
        name: 'Private Notebook',
        groupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4',
        groupName: 'MyGroup'
      }
    }, commandInfo);
    assert.strictEqual(actual, 'Specify either groupId or groupName, but not both');
  });

  it('fails validation if the userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { name: 'Private Notebook', userId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { name: 'Private Notebook', groupId: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid name and groupId specified', async () => {
    const actual = await command.validate({
      options: {
        name: 'Private Notebook',
        groupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid name and groupName specified', async () => {
    const actual = await command.validate({
      options: {
        name: 'Private Notebook',
        groupName: 'MyGroup'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly adds Microsoft OneNote notebook for the currently logged in user (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/onenote/notebooks`) {
        return Promise.resolve({
          "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "self": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "createdDateTime": "2022-10-26T00:05:46Z",
          "displayName": "Private Note",
          "lastModifiedDateTime": "2022-10-26T00:05:46Z",
          "isDefault": false,
          "userRole": "Owner",
          "isShared": false,
          "sectionsUrl": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
          "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
          "createdBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "lastModifiedBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "links": {
            "oneNoteClientUrl": {
              "href": "onenote:https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
            },
            "oneNoteWebUrl": {
              "href": "https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { name: 'Private Notebook', debug: true } });
    assert(loggerLogSpy.calledWith({
      "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "self": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "createdDateTime": "2022-10-26T00:05:46Z",
      "displayName": "Private Note",
      "lastModifiedDateTime": "2022-10-26T00:05:46Z",
      "isDefault": false,
      "userRole": "Owner",
      "isShared": false,
      "sectionsUrl": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
      "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
      "createdBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "links": {
        "oneNoteClientUrl": {
          "href": "onenote:https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
        },
        "oneNoteWebUrl": {
          "href": "https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
        }
      }
    }));
  });

  it('correctly adds Microsoft OneNote notebook for user by id', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks`) {
        return Promise.resolve({
          "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "self": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "createdDateTime": "2022-10-26T00:05:46Z",
          "displayName": "Private Note",
          "lastModifiedDateTime": "2022-10-26T00:05:46Z",
          "isDefault": false,
          "userRole": "Owner",
          "isShared": false,
          "sectionsUrl": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
          "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
          "createdBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "lastModifiedBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "links": {
            "oneNoteClientUrl": {
              "href": "onenote:https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
            },
            "oneNoteWebUrl": {
              "href": "https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { name: 'Private Notebook', userId: 'am917f88-cd36-4048-83c7-6z6608f344f0' } });
    assert(loggerLogSpy.calledWith({
      "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "self": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "createdDateTime": "2022-10-26T00:05:46Z",
      "displayName": "Private Note",
      "lastModifiedDateTime": "2022-10-26T00:05:46Z",
      "isDefault": false,
      "userRole": "Owner",
      "isShared": false,
      "sectionsUrl": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
      "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/users/am917f88-cd36-4048-83c7-6z6608f344f0/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
      "createdBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "links": {
        "oneNoteClientUrl": {
          "href": "onenote:https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
        },
        "oneNoteWebUrl": {
          "href": "https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
        }
      }
    }));
  });

  it('correctly adds Microsoft OneNote notebook for user by name', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/jdoe@contoso.onmicrosoft.com/onenote/notebooks`) {
        return Promise.resolve({
          "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "self": "https://graph.microsoft.com/v1.0/users/jdoe@contoso.onmicrosoft.com/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "createdDateTime": "2022-10-26T00:05:46Z",
          "displayName": "Private Note",
          "lastModifiedDateTime": "2022-10-26T00:05:46Z",
          "isDefault": false,
          "userRole": "Owner",
          "isShared": false,
          "sectionsUrl": "https://graph.microsoft.com/v1.0/users/jdoe@contoso.onmicrosoft.com/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
          "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/users/jdoe@contoso.onmicrosoft.com/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
          "createdBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "lastModifiedBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "links": {
            "oneNoteClientUrl": {
              "href": "onenote:https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
            },
            "oneNoteWebUrl": {
              "href": "https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { name: 'Private Notebook', userName: 'jdoe@contoso.onmicrosoft.com' } });
    assert(loggerLogSpy.calledWith({
      "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "self": "https://graph.microsoft.com/v1.0/users/jdoe@contoso.onmicrosoft.com/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "createdDateTime": "2022-10-26T00:05:46Z",
      "displayName": "Private Note",
      "lastModifiedDateTime": "2022-10-26T00:05:46Z",
      "isDefault": false,
      "userRole": "Owner",
      "isShared": false,
      "sectionsUrl": "https://graph.microsoft.com/v1.0/users/jdoe@contoso.onmicrosoft.com/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
      "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/users/jdoe@contoso.onmicrosoft.com/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
      "createdBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "links": {
        "oneNoteClientUrl": {
          "href": "onenote:https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
        },
        "oneNoteWebUrl": {
          "href": "https://contoso-my.sharepoint.com/personal/jdoe_contoso_onmicrosoft_com/Documents/Notebooks/Private%20Notebook"
        }
      }
    }));
  });

  it('correctly adds Microsoft OneNote notebook in group by id', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks`) {
        return Promise.resolve({
          "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "self": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "createdDateTime": "2022-10-26T00:05:46Z",
          "displayName": "Private Note",
          "lastModifiedDateTime": "2022-10-26T00:05:46Z",
          "isDefault": false,
          "userRole": "Owner",
          "isShared": false,
          "sectionsUrl": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
          "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
          "createdBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "lastModifiedBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "links": {
            "oneNoteClientUrl": {
              "href": "onenote:https://contoso.sharepoint.com/sites/MySite/Shared%20Documents/Notebooks/Private%20Notebook"
            },
            "oneNoteWebUrl": {
              "href": "https://contoso.sharepoint.com/sites/MySite/Shared%20Documents/Notebooks/Private%20Notebook"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { name: 'Private Notebook', groupId: '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4' } });
    assert(loggerLogSpy.calledWith({
      "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "self": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "createdDateTime": "2022-10-26T00:05:46Z",
      "displayName": "Private Note",
      "lastModifiedDateTime": "2022-10-26T00:05:46Z",
      "isDefault": false,
      "userRole": "Owner",
      "isShared": false,
      "sectionsUrl": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
      "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
      "createdBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "links": {
        "oneNoteClientUrl": {
          "href": "onenote:https://contoso.sharepoint.com/sites/MySite/Shared%20Documents/Notebooks/Private%20Notebook"
        },
        "oneNoteWebUrl": {
          "href": "https://contoso.sharepoint.com/sites/MySite/Shared%20Documents/Notebooks/Private%20Notebook"
        }
      }
    }));
  });

  it('handles error when adding Microsoft OneNote notebooks in group by name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { name: 'Private Notebook', groupName: 'MyGroup' } } as any), new CommandError('An error has occurred'));
  });

  it('correctly adds Microsoft OneNote notebook in group by name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4",
              "description": "MyGroup",
              "displayName": "MyGroup"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks`) {
        return Promise.resolve({
          "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "self": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "createdDateTime": "2022-10-26T00:05:46Z",
          "displayName": "Private Note",
          "lastModifiedDateTime": "2022-10-26T00:05:46Z",
          "isDefault": false,
          "userRole": "Owner",
          "isShared": false,
          "sectionsUrl": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
          "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
          "createdBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "lastModifiedBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "links": {
            "oneNoteClientUrl": {
              "href": "onenote:https://contoso.sharepoint.com/sites/testsite/Shared%20Documents/Notebooks/Private%20Notebook"
            },
            "oneNoteWebUrl": {
              "href": "https://contoso.sharepoint.com/sites/testsite/Shared%20Documents/Notebooks/Private%20Notebook"
            }
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { name: 'Private Notebook', groupName: 'MyGroup' } });
    assert(loggerLogSpy.calledWith({
      "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "self": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "createdDateTime": "2022-10-26T00:05:46Z",
      "displayName": "Private Note",
      "lastModifiedDateTime": "2022-10-26T00:05:46Z",
      "isDefault": false,
      "userRole": "Owner",
      "isShared": false,
      "sectionsUrl": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
      "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/groups/233e43d0-dc6a-482e-9b4e-0de7a7bce9b4/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
      "createdBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "links": {
        "oneNoteClientUrl": {
          "href": "onenote:https://contoso.sharepoint.com/sites/testsite/Shared%20Documents/Notebooks/Private%20Notebook"
        },
        "oneNoteWebUrl": {
          "href": "https://contoso.sharepoint.com/sites/testsite/Shared%20Documents/Notebooks/Private%20Notebook"
        }
      }
    }));
  });

  it('handles error when adding Microsoft OneNote notebooks for site', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/sites/`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { name: 'Private Notebook', webUrl: 'https://contoso.sharepoint.com/sites/testsite' } } as any), new CommandError('An error has occurred'));
  });

  it('correctly adds Microsoft OneNote notebook for site', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/v1.0/sites/') > -1) {
        return Promise.resolve({
          "id": "contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2",
          "name": "testsite",
          "webUrl": "https://contoso.sharepoint.com/sites/testsite",
          "displayName": "testsite"
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2/onenote/notebooks`) {
        return Promise.resolve({
          "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "self": "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
          "createdDateTime": "2022-10-26T00:05:46Z",
          "displayName": "Private Note",
          "lastModifiedDateTime": "2022-10-26T00:05:46Z",
          "isDefault": false,
          "userRole": "Owner",
          "isShared": false,
          "sectionsUrl": "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
          "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
          "createdBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "lastModifiedBy": {
            "user": {
              "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
              "displayName": "John Doe"
            }
          },
          "links": {
            "oneNoteClientUrl": {
              "href": "onenote:https://contoso.sharepoint.com/sites/testsite/Shared%20Documents/Notebooks/Private%20Notebook"
            },
            "oneNoteWebUrl": {
              "href": "https://contoso.sharepoint.com/sites/testsite/Shared%20Documents/Notebooks/Private%20Notebook"
            }
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { name: 'Private Notebook', webUrl: 'https://contoso.sharepoint.com/sites/testsite' } });
    assert(loggerLogSpy.calledWith({
      "id": "1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "self": "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0",
      "createdDateTime": "2022-10-26T00:05:46Z",
      "displayName": "Private Note",
      "lastModifiedDateTime": "2022-10-26T00:05:46Z",
      "isDefault": false,
      "userRole": "Owner",
      "isShared": false,
      "sectionsUrl": "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sections",
      "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,c2ceff0c-063b-45b3-a9ec-3a7f8e67547f,4aef2b1f-7a54-4f54-be16-167abba63cf2/onenote/notebooks/1-558ac4dh-3c0a-4123-bc46-e4d1c22256f0/sectionGroups",
      "createdBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "am917f88-cd36-4048-83c7-6z6608f344f0",
          "displayName": "John Doe"
        }
      },
      "links": {
        "oneNoteClientUrl": {
          "href": "onenote:https://contoso.sharepoint.com/sites/testsite/Shared%20Documents/Notebooks/Private%20Notebook"
        },
        "oneNoteWebUrl": {
          "href": "https://contoso.sharepoint.com/sites/testsite/Shared%20Documents/Notebooks/Private%20Notebook"
        }
      }
    }));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
