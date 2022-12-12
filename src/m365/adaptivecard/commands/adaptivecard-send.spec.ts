import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
import { Cli } from '../../../cli/Cli';
import { CommandInfo } from '../../../cli/CommandInfo';
import { Logger } from '../../../cli/Logger';
import Command, { CommandError } from '../../../Command';
import request from '../../../request';
import { pid } from '../../../utils/pid';
import { sinonUtil } from '../../../utils/sinonUtil';
import commands from '../commands';
const command: Command = require('./adaptivecard-send');
// required to avoid tests from timing out due to dynamic imports
import 'adaptivecards-templating';

describe(commands.SEND, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SEND), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  describe('send card to Teams', () => {
    it('sends card with just title', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });

      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4'
        }
      });
    });

    it('sends card with just title (debug)', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });

      await command.action(logger, {
        options: {
          debug: true,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4'
        }
      });
    });

    it('sends card with just description', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });

      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          description: 'New release of CLI for Microsoft 365'
        }
      });
    });

    it('sends card with title and description', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365'
        }
      });
    });

    it('sends card with title, description and image', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          imageUrl: 'https://contoso.com/image.gif'
        }
      });
    });

    it('sends card with title, description and action', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          actionUrl: 'https://aka.ms/cli-m365'
        }
      });
    });

    it('sends card with title, description, image and action', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          imageUrl: 'https://contoso.com/image.gif',
          actionUrl: 'https://aka.ms/cli-m365'
        }
      });
    });

    it('sends card with title, description, action and unknown options', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          actionUrl: 'https://aka.ms/cli-m365',
          Version: 'v3.4.0',
          ReleaseNotes: 'https://pnp.github.io/cli-microsoft365/about/release-notes/#v340'
        }
      });
    });

    it('sends custom card without any data', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${viewUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}'
        }
      });
    });

    it('sends custom card with just title merged', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${viewUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}',
          title: 'CLI for Microsoft 365 v3.4'
        }
      });
    });

    it('sends custom card with all known options merged', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${actionUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}',
          title: 'CLI for Microsoft 365 v3.4',
          description: 'New release of CLI for Microsoft 365',
          imageUrl: 'https://contoso.com/image.gif',
          actionUrl: 'https://aka.ms/cli-m365'
        }
      });
    });

    it('sends custom card with unknown option merged', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${Title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${viewUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}',
          Title: 'CLI for Microsoft 365 v3.4'
        }
      });
    });

    it('sends custom card with cardData', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve(1);
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          card: '{"type":"AdaptiveCard","body":[{"type":"TextBlock","size":"Medium","weight":"Bolder","text":"${title}"},{"type":"TextBlock","text":"${description}","wrap":true},{"type":"FactSet","facts":[{"$data":"${properties}","title":"${key}:","value":"${value}"}]}],"actions":[{"type":"Action.OpenUrl","title":"View","url":"${viewUrl}"}],"$schema":"http://adaptivecards.io/schemas/adaptive-card.json","version":"1.2"}',
          cardData: '{"title":"Publish Adaptive Card Schema","description":"Now that we have defined the main rules and features of the format, we need to produce a schema and publish it to GitHub. The schema will be the starting point of our reference documentation.","creator":{"name":"Matt Hidinger","profileImage":"https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg"},"createdUtc":"2017-02-14T06:08:39Z","viewUrl":"https://adaptivecards.io","properties":[{"key":"Board","value":"Adaptive Cards"},{"key":"List","value":"Backlog"},{"key":"Assigned to","value":"Matt Hidinger"},{"key":"Due date","value":"Not set"}]}'
        }
      });
    });

    it('correctly handles error when sending card to Teams', async () => {
      sinon.stub(request, 'post').callsFake(_ => Promise.resolve('Webhook message delivery failed with error: Microsoft Teams endpoint returned HTTP error 400 with ContextId MS-CV=Qn6afVIGzEq'));
      await assert.rejects(command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4'
        }
      }), new CommandError('Webhook message delivery failed with error: Microsoft Teams endpoint returned HTTP error 400 with ContextId MS-CV=Qn6afVIGzEq'));
    });
  });

  describe('send card to a URL', () => {
    it('sends card with just title', async () => {
      sinon.stub(request, 'post').callsFake(opts => {
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
          return Promise.resolve('OK');
        }

        return Promise.reject(`Invalid data: ${JSON.stringify(opts.data)}`);
      });
      await command.action(logger, {
        options: {
          debug: false,
          url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547',
          title: 'CLI for Microsoft 365 v3.4'
        }
      });
    });
  });

  it(`passes validation if the title is specified`, async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', title: 'Lorem' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it(`fails validation if the specified card is not a valid JSON string`, async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', card: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if the specified card is a valid JSON string`, async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', card: '{}' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it(`fails validation if specified cardData without card`, async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', cardData: '{}' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`fails validation if specified cardData is not a valid JSON string`, async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', card: '{}', cardData: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if the specified cardData is a valid JSON string`, async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.webhook.office.com/webhookb2/892e8ed3-997c-4b6e-8f8a-7f32728a8a87@f7322380-f203-42ff-93e8-66e266f6d2e4/IncomingWebhook/fcc6565ec7a944928bd43d6fc193b258/4f0482d4-b147-4f67-8a61-11f0a5019547', card: '{}', cardData: '{}' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [{ options: ['title', 'card'] }]);
  });

  it('supports specifying unknown options', () => {
    assert.strictEqual(command.allowUnknownOptions(), true);
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