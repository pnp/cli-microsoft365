import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../Auth.js';
import { cli } from '../../../cli/cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { CommandError } from '../../../Command.js';
import request from '../../../request.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import { sinonUtil } from '../../../utils/sinonUtil.js';
import commands from '../commands.js';
import command, { options } from './adaptivecard-send.js';
// required to avoid tests from timing out due to dynamic imports
import 'adaptivecards-templating';
import { settingsNames } from '../../../settingsNames.js';

describe(commands.SEND, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SEND);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  describe('send card to Teams', () => {
    it('sends card with just title', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  }
                ]
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });

      await command.action(logger, {
        options: {
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4'
        }
      });
    });

    it('sends card with just title (debug)', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  }
                ]
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });

      await command.action(logger, {
        options: commandOptionsSchema.parse({
          debug: true,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4'
        })
      });
    });

    it('sends card with just description', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2",
                "body": [
                  {
                    "type": "TextBlock",
                    "text": "New release of CLI for Microsoft 365",
                    "wrap": true
                  }
                ]
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });

      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          description: 'New release of CLI for Microsoft 365'
        })
      });
    });

    it('sends card with title and description', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  },
                  {
                    "type": "TextBlock",
                    "text": "New release of CLI for Microsoft 365",
                    "wrap": true
                  }
                ]
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365'
        })
      });
    });

    it('sends card with title, description and image', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  },
                  {
                    "type": "Image",
                    "url": "https://contoso.com/image.gif",
                    "size": "Stretch"
                  },
                  {
                    "type": "TextBlock",
                    "text": "New release of CLI for Microsoft 365",
                    "wrap": true
                  }
                ]
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          imageUrl: 'https://contoso.com/image.gif'
        })
      });
    });

    it('sends card with title, description and action', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  },
                  {
                    "type": "TextBlock",
                    "text": "New release of CLI for Microsoft 365",
                    "wrap": true
                  }
                ],
                "actions": [
                  {
                    "type": "Action.OpenUrl",
                    "title": "View",
                    "url": "https://aka.ms/cli-m365"
                  }
                ]
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          actionUrl: 'https://aka.ms/cli-m365'
        })
      });
    });

    it('sends card with title, description, image and action', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  },
                  {
                    "type": "Image",
                    "url": "https://contoso.com/image.gif",
                    "size": "Stretch"
                  },
                  {
                    "type": "TextBlock",
                    "text": "New release of CLI for Microsoft 365",
                    "wrap": true
                  }
                ],
                "actions": [
                  {
                    "type": "Action.OpenUrl",
                    "title": "View",
                    "url": "https://aka.ms/cli-m365"
                  }
                ]
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          imageUrl: 'https://contoso.com/image.gif',
          actionUrl: 'https://aka.ms/cli-m365'
        })
      });
    });

    it('sends card with title, description, action and unknown options', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  },
                  {
                    "type": "TextBlock",
                    "text": "New release of CLI for Microsoft 365",
                    "wrap": true
                  },
                  {
                    "type": "FactSet",
                    "facts": [
                      {
                        "title": "Version:",
                        "value": "v3.4.0"
                      },
                      {
                        "title": "ReleaseNotes:",
                        "value": "https://pnp.github.io/cli-microsoft365/about/release-notes/#v340"
                      }
                    ]
                  }
                ],
                "actions": [
                  {
                    "type": "Action.OpenUrl",
                    "title": "View",
                    "url": "https://aka.ms/cli-m365"
                  }
                ]
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          actionUrl: 'https://aka.ms/cli-m365',
          Version: 'v3.4.0',
          ReleaseNotes: 'https://pnp.github.io/cli-microsoft365/about/release-notes/#v340'
        })
      });
    });

    it('sends custom card without any data', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "${title}"
                  },
                  {
                    "type": "TextBlock",
                    "text": "${description}",
                    "wrap": true
                  },
                  {
                    "type": "FactSet",
                    "facts": [
                      {
                        "$data": "${properties}",
                        "title": "${key}:",
                        "value": "${value}"
                      }
                    ]
                  }
                ],
                "actions": [
                  {
                    "type": "Action.OpenUrl",
                    "title": "View",
                    "url": "${viewUrl}"
                  }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${viewUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}'
        })
      });
    });

    it('sends custom card with just title merged', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  },
                  {
                    "type": "TextBlock",
                    "text": "${description}",
                    "wrap": true
                  },
                  {
                    "type": "FactSet",
                    "facts": [
                      {
                        "title": "${key}:",
                        "value": "${value}"
                      }
                    ]
                  }
                ],
                "actions": [
                  {
                    "type": "Action.OpenUrl",
                    "title": "View",
                    "url": "${viewUrl}"
                  }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });

      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${viewUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}',
          title: 'CLI for Microsoft 365 v3.4'
        })
      });
    });

    it('sends custom card with all known options merged', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  },
                  {
                    "type": "TextBlock",
                    "text": "New release of CLI for Microsoft 365",
                    "wrap": true
                  },
                  {
                    "type": "FactSet",
                    "facts": [
                      {
                        "title": "${key}:",
                        "value": "${value}"
                      }
                    ]
                  }
                ],
                "actions": [
                  {
                    "type": "Action.OpenUrl",
                    "title": "View",
                    "url": "https://aka.ms/cli-m365"
                  }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });

      // For this test we need the base schema without the refinement
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${actionUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          imageUrl: 'https://contoso.com/image.gif',
          actionUrl: 'https://aka.ms/cli-m365'
        })
      });
    });

    it('sends custom card with unknown option merged', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  },
                  {
                    "type": "TextBlock",
                    "text": "${description}",
                    "wrap": true
                  },
                  {
                    "type": "FactSet",
                    "facts": [
                      {
                        "title": "${key}:",
                        "value": "${value}"
                      }
                    ]
                  }
                ],
                "actions": [
                  {
                    "type": "Action.OpenUrl",
                    "title": "View",
                    "url": "${viewUrl}"
                  }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${Title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${viewUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}',
          Title: 'CLI for Microsoft 365 v3.4'
        })
      });
    });

    it('sends custom card with cardData', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "Publish Adaptive Card Schema"
                  },
                  {
                    "type": "TextBlock",
                    "text": "Now that we have defined the main rules and features of the format, we need to produce a schema and publish it to GitHub. The schema will be the starting point of our reference documentation.",
                    "wrap": true
                  },
                  {
                    "type": "FactSet",
                    "facts": [
                      {
                        "title": "Board:",
                        "value": "Adaptive Cards"
                      },
                      {
                        "title": "List:",
                        "value": "Backlog"
                      },
                      {
                        "title": "Assigned to:",
                        "value": "Matt Hidinger"
                      },
                      {
                        "title": "Due date:",
                        "value": "Not set"
                      }
                    ]
                  }
                ],
                "actions": [
                  {
                    "type": "Action.OpenUrl",
                    "title": "View",
                    "url": "https://adaptivecards.io"
                  }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2"
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${viewUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}',
          cardData: '{"title":"Publish Adaptive Card Schema","description":"Now that we have defined the main rules and features of the format, we need to produce a schema and publish it to GitHub. The schema will be the starting point of our reference documentation.","creator":{"name":"Matt Hidinger","profileImage":"https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg"},"createdUtc":"2017-02-14T06:08:39Z","viewUrl":"https://adaptivecards.io","properties":[{"key":"Board","value":"Adaptive Cards"},{"key":"List","value":"Backlog"},{"key":"Assigned to","value":"Matt Hidinger"},{"key":"Due date","value":"Not set"}]}'
        })
      });
    });

    it('correctly handles error when sending card to Teams', async () => {
      sinon.stub(request, 'post').resolves('Webhook message delivery failed with error: Microsoft Teams endpoint returned HTTP error 400 with ContextId MS-CV=Qn6afVIGzEq');
      await assert.rejects(command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4'
        })
      }), new CommandError('Webhook message delivery failed with error: Microsoft Teams endpoint returned HTTP error 400 with ContextId MS-CV=Qn6afVIGzEq'));
    });
  });

  describe('send card to a URL', () => {
    it('sends card with just title', async () => {
      sinon.stub(request, 'post').callsFake(async opts => {
        if (JSON.stringify(opts.data) === JSON.stringify({
          "type": "message",
          "attachments": [
            {
              "contentType": "application/vnd.microsoft.card.adaptive",
              "content": {
                "type": "AdaptiveCard",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.2",
                "body": [
                  {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "CLI for Microsoft 365 v3.4"
                  }
                ]
              }
            }
          ]
        })) {
          return '1';
        }

        throw `Invalid data: ${JSON.stringify(opts.data)}`;
      });
      await command.action(logger, {
        options: commandOptionsSchema.parse({
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4'
        })
      });
    });
  });

  it(`passes validation if the title is specified`, () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', title: 'Lorem' });
    assert.strictEqual(actual.success, true);
  });

  it(`fails validation if the specified card is not a valid JSON string`, () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', card: 'abc' });
    assert.notStrictEqual(actual.success, true);
  });

  it(`passes validation if the specified card is a valid JSON string`, () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', card: '{}' });
    assert.strictEqual(actual.success, true);
  });

  it(`fails validation if specified cardData without card`, () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', cardData: '{}' });
    assert.strictEqual(actual.success, false);
  });

  it(`fails validation if specified cardData is not a valid JSON string`, () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', card: '{}', cardData: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it(`passes validation if the specified cardData is a valid JSON string`, () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', card: '{}', cardData: '{}' });
    assert.strictEqual(actual.success, true);
  });
});