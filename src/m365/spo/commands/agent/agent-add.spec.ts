/* eslint-disable camelcase */
import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './agent-add.js';

describe(commands.AGENT_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    sinon.stub(spo, 'ensureFolder').resolves();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.AGENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a SharePoint agent with required options', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='/sites/test/SiteAssets/Copilots')/Files/AddUsingPath(DecodedUrl='Test%20Agent.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)"
      ) {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/web/lists/EnsureSiteAssetsLibrary()') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/search/postquery') {
        return {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: [
                  {
                    Cells: [
                      { Key: "contentclass", Value: "STS_ListItem_DocumentLibrary" },
                      { Key: "Title", Value: "Test Document" },
                      { Key: "Path", Value: "https://contoso.sharepoint.com/sites/test/Shared Documents/Test Document.docx" },
                      { Key: "SiteName", Value: "Test Site" },
                      { Key: "SiteTitle", Value: "Test Site" },
                      { Key: "ListID", Value: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a" },
                      { Key: "ListItemID", Value: "a7c6b5d4-e3f2-1a09-b8c7-6e5d4c3b2a1f" },
                      { Key: "SiteID", Value: "f1e2d3c4-b5a6-7890-1234-56789abcdef0" },
                      { Key: "WebId", Value: "123e4567-e89b-12d3-a456-426614174000" },
                      { Key: "UniqueID", Value: "{0f1e2d3c-4b5a-6789-0123-456789abcdef}" },
                      { Key: "IsDocument", Value: "true" },
                      { Key: "IsContainer", Value: "false" }
                    ]
                  }
                ]
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Test Agent',
        agentInstructions: 'You are a helpful test agent',
        welcomeMessage: 'Hello! I am your test agent.',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test',
        description: 'A test agent'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data,
      {
        schemaVersion: "0.2.0",
        customCopilotConfig: {
          conversationStarters: {
            conversationStarterList: [],
            welcomeMessage: {
              text: "Hello! I am your test agent."
            }
          },
          gptDefinition: {
            name: "Test Agent",
            description: "A test agent",
            instructions: "You are a helpful test agent",
            capabilities: [
              {
                name: "OneDriveAndSharePoint",
                items_by_sharepoint_ids: [
                  {
                    url: "https://contoso.sharepoint.com/sites/test",
                    name: "Test Document",
                    site_id: "f1e2d3c4-b5a6-7890-1234-56789abcdef0",
                    web_id: "123e4567-e89b-12d3-a456-426614174000",
                    list_id: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a",
                    unique_id: "0f1e2d3c-4b5a-6789-0123-456789abcdef",
                    type: "File"
                  }
                ],
                items_by_url: []
              }
            ]
          },
          icon: "data:image/png;base64,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"
        }
      }
    );
  });

  it('creates a SharePoint agent with all options including optional ones', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='/sites/test/SiteAssets/Copilots')/Files/AddUsingPath(DecodedUrl='Complete%20Agent.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)"
      ) {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/web/lists/EnsureSiteAssetsLibrary()') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/search/postquery') {
        return {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: [
                  {
                    Cells: [
                      { Key: "contentclass", Value: "STS_ListItem_DocumentLibrary" },
                      { Key: "Title", Value: "Test Document" },
                      { Key: "Path", Value: "https://contoso.sharepoint.com/sites/test/Shared Documents/Test Document.docx" },
                      { Key: "SiteName", Value: "Test Site" },
                      { Key: "SiteTitle", Value: "Test Site" },
                      { Key: "ListID", Value: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a" },
                      { Key: "ListItemID", Value: "a7c6b5d4-e3f2-1a09-b8c7-6e5d4c3b2a1f" },
                      { Key: "SiteID", Value: "f1e2d3c4-b5a6-7890-1234-56789abcdef0" },
                      { Key: "WebId", Value: "123e4567-e89b-12d3-a456-426614174000" },
                      { Key: "UniqueID", Value: "{0f1e2d3c-4b5a-6789-0123-456789abcdef}" },
                      { Key: "IsDocument", Value: "true" },
                      { Key: "IsContainer", Value: "false" }
                    ]
                  }
                ]
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Complete Agent',
        agentInstructions: 'You are a comprehensive test agent',
        welcomeMessage: 'Welcome to the comprehensive test agent',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test/Shared Documents/Test Document.docx,https://contoso.sharepoint.com/sites/test/Shared Documents/Test Document.docx',
        description: 'A comprehensive test agent',
        icon: 'https://contoso.sharepoint.com/sites/test/SiteAssets/agent-icon.png',
        conversationStarters: 'What can you help me with?,Show me recent documents,Help me with my tasks',
        verbose: true
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data,
      {
        schemaVersion: "0.2.0",
        customCopilotConfig: {
          conversationStarters: {
            conversationStarterList: [
              {
                text: "What can you help me with?"
              },
              {
                text: "Show me recent documents"
              },
              {
                text: "Help me with my tasks"
              }
            ],
            welcomeMessage: {
              text: "Welcome to the comprehensive test agent"
            }
          },
          gptDefinition: {
            name: "Complete Agent",
            description: "A comprehensive test agent",
            instructions: "You are a comprehensive test agent",
            capabilities: [
              {
                name: "OneDriveAndSharePoint",
                items_by_sharepoint_ids: [
                  {
                    url: "https://contoso.sharepoint.com/sites/test/Shared Documents/Test Document.docx",
                    name: "Test Document",
                    site_id: "f1e2d3c4-b5a6-7890-1234-56789abcdef0",
                    web_id: "123e4567-e89b-12d3-a456-426614174000",
                    list_id: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a",
                    unique_id: "0f1e2d3c-4b5a-6789-0123-456789abcdef",
                    type: "File"
                  },
                  {
                    url: "https://contoso.sharepoint.com/sites/test/Shared Documents/Test Document.docx",
                    name: "Test Document",
                    site_id: "f1e2d3c4-b5a6-7890-1234-56789abcdef0",
                    web_id: "123e4567-e89b-12d3-a456-426614174000",
                    list_id: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a",
                    unique_id: "0f1e2d3c-4b5a-6789-0123-456789abcdef",
                    type: "File"
                  }
                ],
                items_by_url: [
                ]
              }
            ]
          },
          icon: "https://contoso.sharepoint.com/sites/test/SiteAssets/agent-icon.png"
        }
      }
    );
  });

  it('creates a SharePoint agent with not found resource in search', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='/sites/test/SiteAssets/Copilots')/Files/AddUsingPath(DecodedUrl='Complete%20Agent.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)"
      ) {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/web/lists/EnsureSiteAssetsLibrary()') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/search/postquery') {
        return {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: [
                ]
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Complete Agent',
        agentInstructions: 'You are a comprehensive test agent',
        welcomeMessage: 'Welcome to the comprehensive test agent',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test/Shared Documents/Test Document.docx',
        description: 'A comprehensive test agent',
        icon: 'https://contoso.sharepoint.com/sites/test/SiteAssets/agent-icon.png',
        conversationStarters: 'What can you help me with?,Show me recent documents,Help me with my tasks'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data,
      {
        schemaVersion: "0.2.0",
        customCopilotConfig: {
          conversationStarters: {
            conversationStarterList: [
              {
                text: "What can you help me with?"
              },
              {
                text: "Show me recent documents"
              },
              {
                text: "Help me with my tasks"
              }
            ],
            welcomeMessage: {
              text: "Welcome to the comprehensive test agent"
            }
          },
          gptDefinition: {
            name: "Complete Agent",
            description: "A comprehensive test agent",
            instructions: "You are a comprehensive test agent",
            capabilities: [
              {
                name: "OneDriveAndSharePoint",
                items_by_sharepoint_ids: [
                ],
                items_by_url: [
                ]
              }
            ]
          },
          icon: "https://contoso.sharepoint.com/sites/test/SiteAssets/agent-icon.png"
        }
      }
    );
  });

  it('creates a SharePoint agent with Site resource', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='/sites/test/SiteAssets/Copilots')/Files/AddUsingPath(DecodedUrl='Test%20Agent.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)"
      ) {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/web/lists/EnsureSiteAssetsLibrary()') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/search/postquery') {
        return {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: [
                  {
                    Cells: [
                      { Key: "contentclass", Value: "STS_Site" },
                      { Key: "Title", Value: "Test Site" },
                      { Key: "Path", Value: "https://contoso.sharepoint.com/sites/test" },
                      { Key: "SiteName", Value: "Test Site" },
                      { Key: "SiteTitle", Value: "Test Site" },
                      { Key: "SiteID", Value: "f1e2d3c4-b5a6-7890-1234-56789abcdef0" },
                      { Key: "WebId", Value: "123e4567-e89b-12d3-a456-426614174000" },
                      { Key: "UniqueID", Value: "{0f1e2d3c-4b5a-6789-0123-456789abcdef}" },
                      { Key: "IsDocument", Value: "false" },
                      { Key: "IsContainer", Value: "false" }
                    ]
                  }
                ]
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Test Agent',
        agentInstructions: 'You are a helpful test agent',
        welcomeMessage: 'Hello! I am your test agent.',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test',
        description: 'A test agent'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data,
      {
        schemaVersion: "0.2.0",
        customCopilotConfig: {
          conversationStarters: {
            conversationStarterList: [],
            welcomeMessage: {
              text: "Hello! I am your test agent."
            }
          },
          gptDefinition: {
            name: "Test Agent",
            description: "A test agent",
            instructions: "You are a helpful test agent",
            capabilities: [
              {
                name: "OneDriveAndSharePoint",
                items_by_sharepoint_ids: [
                ],
                items_by_url: [
                  {
                    url: "https://contoso.sharepoint.com/sites/test",
                    name: "Test Site",
                    site_id: "f1e2d3c4-b5a6-7890-1234-56789abcdef0",
                    web_id: "123e4567-e89b-12d3-a456-426614174000",
                    list_id: "",
                    unique_id: "0f1e2d3c-4b5a-6789-0123-456789abcdef",
                    type: "Site"
                  }
                ]
              }
            ]
          },
          icon: "data:image/png;base64,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"
        }
      }
    );
  });

  it('creates a SharePoint agent with Subsite resource', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='/sites/test/SiteAssets/Copilots')/Files/AddUsingPath(DecodedUrl='Test%20Agent.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)"
      ) {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/web/lists/EnsureSiteAssetsLibrary()') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/search/postquery') {
        return {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: [
                  {
                    Cells: [
                      { Key: "contentclass", Value: "STS_Web" },
                      { Key: "Title", Value: "Test Site" },
                      { Key: "Path", Value: "https://contoso.sharepoint.com/sites/test/subsite" },
                      { Key: "SiteName", Value: "Test Site" },
                      { Key: "SiteTitle", Value: "Test Site" },
                      { Key: "SiteID", Value: "f1e2d3c4-b5a6-7890-1234-56789abcdef0" },
                      { Key: "WebId", Value: "123e4567-e89b-12d3-a456-426614174000" },
                      { Key: "UniqueID", Value: "{0f1e2d3c-4b5a-6789-0123-456789abcdef}" },
                      { Key: "IsDocument", Value: "false" },
                      { Key: "IsContainer", Value: "false" }
                    ]
                  }
                ]
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Test Agent',
        agentInstructions: 'You are a helpful test agent',
        welcomeMessage: 'Hello! I am your test agent.',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test/subsite',
        description: 'A test agent'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data,
      {
        schemaVersion: "0.2.0",
        customCopilotConfig: {
          conversationStarters: {
            conversationStarterList: [],
            welcomeMessage: {
              text: "Hello! I am your test agent."
            }
          },
          gptDefinition: {
            name: "Test Agent",
            description: "A test agent",
            instructions: "You are a helpful test agent",
            capabilities: [
              {
                name: "OneDriveAndSharePoint",
                items_by_sharepoint_ids: [
                ],
                items_by_url: [
                  {
                    url: "https://contoso.sharepoint.com/sites/test/subsite",
                    name: "Test Site",
                    site_id: "f1e2d3c4-b5a6-7890-1234-56789abcdef0",
                    web_id: "123e4567-e89b-12d3-a456-426614174000",
                    list_id: "",
                    unique_id: "0f1e2d3c-4b5a-6789-0123-456789abcdef",
                    type: "Site"
                  }
                ]
              }
            ]
          },
          icon: "data:image/png;base64,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"
        }
      }
    );
  });

  it('creates a SharePoint agent with Document Library resource', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='/sites/test/SiteAssets/Copilots')/Files/AddUsingPath(DecodedUrl='Test%20Agent.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)"
      ) {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/web/lists/EnsureSiteAssetsLibrary()') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/search/postquery') {
        return {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: [
                  {
                    Cells: [
                      { Key: "contentclass", Value: "STS_List_DocumentLibrary" },
                      { Key: "Title", Value: "Test Library" },
                      { Key: "Path", Value: "https://contoso.sharepoint.com/sites/test/documents" },
                      { Key: "SiteName", Value: "Test Site" },
                      { Key: "SiteTitle", Value: "Test Site" },
                      { Key: "SiteID", Value: "f1e2d3c4-b5a6-7890-1234-56789abcdef0" },
                      { Key: "WebId", Value: "123e4567-e89b-12d3-a456-426614174000" },
                      { Key: "ListID", Value: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a" },
                      { Key: "UniqueID", Value: "{0f1e2d3c-4b5a-6789-0123-456789abcdef}" },
                      { Key: "IsDocument", Value: "false" },
                      { Key: "IsContainer", Value: "false" }
                    ]
                  }
                ]
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Test Agent',
        agentInstructions: 'You are a helpful test agent',
        welcomeMessage: 'Hello! I am your test agent.',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test/documents',
        description: 'A test agent'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data,
      {
        schemaVersion: "0.2.0",
        customCopilotConfig: {
          conversationStarters: {
            conversationStarterList: [],
            welcomeMessage: {
              text: "Hello! I am your test agent."
            }
          },
          gptDefinition: {
            name: "Test Agent",
            description: "A test agent",
            instructions: "You are a helpful test agent",
            capabilities: [
              {
                name: "OneDriveAndSharePoint",
                items_by_sharepoint_ids: [
                ],
                items_by_url: [
                  {
                    url: "https://contoso.sharepoint.com/sites/test/documents",
                    name: "Test Library",
                    site_id: "f1e2d3c4-b5a6-7890-1234-56789abcdef0",
                    web_id: "123e4567-e89b-12d3-a456-426614174000",
                    list_id: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a",
                    unique_id: "0f1e2d3c-4b5a-6789-0123-456789abcdef",
                    type: "List"
                  }
                ]
              }
            ]
          },
          icon: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAYAAACOEfKtAAAIJklEQVR4nO2bCWwUVRjHv9m73RYEBEQR5RLKpSgKqHggtBQRBAQFEY94GzUcVlBDQIkhGhUFJCqCBFDuUzkrN4qCii1GG4UKqFwCgba7OzO7M37fN3a729mFbh+IJO+XkG135r33vf981xtAOTGgnwmSauM43wZc6EgBBZECCiIFFEQKKIgUUBApoCBSQEGkgIJIAQWRAgoiBRRECiiIFFAQKaAgUkBBpICCSAEFkQIKIgUURAooiBRQECmgIFJAQaSAgkgBBZECCiIFFMQuoOMC0NTlAsXjOd9WMK7KX6Q/+TQYhw+B+vkKMFX1fNiUGEUBb7ds8OTmgvOyhvy7cewY6Fs3Q2jRQjBDoaTj3FdfA+7ON4Kz0RUATieU5I2w3eaoXx8yXnol6fL6zp0QnDXT9r1NQMXnA9+9g8DTPQdC8+eCtnEDgGGksNNzQ/rjT4Lnjm5glpSAtmUz7kgHV5u24O3TF5wtsqB03BiASCRujJKRAf7nh4ELBSTMU6cggs6RCEedi8HR4FJrrwn2q9SokXCcTUBt+9fgbJkFjtq12Ru9Pe+E0OxZoO/6IeVNny1cWa1YPOPwYSh5eTQKcdK64HZD5rjXwNWyJbg7dgL9q20Vg9DT/C++BK4WLSBcWADBObMgsndv0jUctWrxZ2jJIgjNm1tl22wJj4woee4Z9L55YAaD7PZ+dO2MV8aA88rGVZ7Ymt0Rl1MddeuBs1nzqLGx9zkbXg7Opk1B8WfYpnG1b8+fav7aCvHYWB20zZuse1q0jBvjRcFJPP3776D09fGnFY9Q/rWJ0kIq2DyQoHwSWjgf1HVrwNd/AOae7uBqdzVkTmjL4ROa+xku9PcZJ6ecQuMCUyaDN7cnOJs0iV6jjQUmvcvX04Y+hCFUx7qA4aNtWA+BaR9hSIb5K3XFctDy8zF8T9ltRREZ9LhYPLgehXSQ54nYxlXGUau2NV+KAjpHtc4am/QqFpHwD9+DtnUrKDVrojc2Yi/0ZGeD4vXhU93DXpAMzy23YnK+BNwdOkBk/z7QUZhI8V5wXtIAnI0bgxs9y5uTC+GiX0DfiNcOHEBPbGh5E84b/uVnayJNA7OsDCAcjl8AC0Ta4Ae4AKhYSIy//rLEqFsX0u4bzOPDu3eDD/OkN6cHuK+9jqs3rWOz9fauvD8a4+lyC3i73gGutu1AwQdqHEqcN9mEVP6NNHkQGexq145/p6TMnpq/zr45hMKePExdtRKCMz6umOfyRpD51jv8s7p6FQSnT4tec7e/FvyjX2bBS0YOt9vQHFNA7Toc6p4uXcDVug1omzawl0fn6HA9+PNGgXHkiOXZKDR1FFQg6Wd9x7dQ9tabccUiY+yr4GrVOuG+KeoCk9/DDdulShjCyaA8Ujp+nBV29w9BL0JBH3mUwzP46RzQv9mecBwZHDfPgf1gnDjOYUMCxlLudZQvE+Hr1ZtbklibgjM/ibuHKip/1qvHOTKEBcQ4cYKjxz8yD9zX3wDe7tmgrlkdHUO5zwwEQF2+FDQqRhj25LG+IUPZIyNFRaCuXQ2VqVbXHC74EUpG5UHgvYmWEFj+/SNe4GpZ9UksjzUDZfHfU76iJ52koQ/hBsvemACB9yfzA6OoyJzwBqaYiypu8nqtJTB8A1MmsXg89e/FmFs/4J89XbvFzUv5+OTDQyG0eBGHrHH0KAsc/NiKDk92TkJ7qn3sUGpgTsRc5ci0+iNTxzwVDKQ+UYr/ySKyZw82tTu4P6UwVJcvQ0+rD75+/WOMsz4ot1YOu3BhIYcz5Ts60cTbYjdG27YFB+mcmxOdflIKYbYtLQ28GEbeu3pbOcU0/q3Mn/JTSxkl9SGxUIh6e/eJ834KRZ46Pd0+AD2c2jMqigoKaCbI3XFgMTMDQauRJgGxoMVSdQFxMWpnqK2hxYnw7kIIzp5lVePqUgUPTHvsCT6OBT6Yyk1x/Hj7BOU9n6v5VbZrJKojM5OrevlRlXPd0AfZq9WlS2z3K34/mJrKwlfmzALSWfLGm7AtGMQtCRuIRSA0Zzb3cv8FJuYwKgjUFlUW0N2xo2XTb79Gv6MHGvnzD2zam3FV17EVK8eD52nqGfUfd0XF5/YJ87g3uwdomPdihaL2hu4PFxQk7CdPKyD1QVxtmzTl3ykZ0wlFw56tKs3p2UJdu4aTuOfW2zgSuNpj6PFZuEcubziEuTAKChOcMZ0b+fRhI9CrFmNbtB+PfFmYfnqx56kLF0RvN44eAW39l3xcpHZGXfkFn7mdmBZ8ve7ivaqLF9gNgyQCcnuCwlG7wvbgyURdsYxPBEnfepxD6PhWOnYMns2fAs9NN/Ofcsh7AlOnYOU8GDeGOoWyiW9DOoa/D5vqckgs6hnJQ2MJYLU1sVjQG5/0Z56tuB+dJjjtQyxIRQltszXSvnsGgm/AQA5dMCL4ZNZDaMG8aCuQCtzEYktBx6PKr8bo9EChYRw6zOtUWKRgW9SAvcg4eBAqQ9fo3ExtDr12i+zblzAPRqfD9alboFxGqSBMoX6a6KH87mrajF9UmCdPWvefptDYBPQPHwnuTp1B/24n57nIH/Zjj6QCWwhHiou5gQz/tPt82HPBYROQ3odJqk7KjXRVobfBvrv74gru6k+CuZH+asE4fvzsGXaWOacC8qt0QQG1bXiw/x8LmNLrLImdC+DvMP/fSAEFkQIKIgUURAooiBRQECmgIFJAQaSAgkgBBZECCiIFFEQKKIgUUBApoCBSQEGkgIJIAQWRAgoiBRRECiiIFFAQKaAgUkBBpICCSAEFkQIKIgUU5B91HS13TtrWPgAAAABJRU5ErkJggg=="
        }
      }
    );
  });

  it('creates a SharePoint agent with Folder resource', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='/sites/test/SiteAssets/Copilots')/Files/AddUsingPath(DecodedUrl='Test%20Agent.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)"
      ) {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/web/lists/EnsureSiteAssetsLibrary()') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/search/postquery') {
        return {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: [
                  {
                    Cells: [
                      { Key: "contentclass", Value: "STS_ListItem_DocumentLibrary" },
                      { Key: "Title", Value: "Test Folder" },
                      { Key: "Path", Value: "https://contoso.sharepoint.com/sites/test/documents/folder" },
                      { Key: "SiteName", Value: "Test Site" },
                      { Key: "SiteTitle", Value: "Test Site" },
                      { Key: "SiteID", Value: "f1e2d3c4-b5a6-7890-1234-56789abcdef0" },
                      { Key: "WebId", Value: "123e4567-e89b-12d3-a456-426614174000" },
                      { Key: "ListID", Value: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a" },
                      { Key: "UniqueID", Value: "{0f1e2d3c-4b5a-6789-0123-456789abcdef}" },
                      { Key: "IsDocument", Value: "false" },
                      { Key: "IsContainer", Value: "true" }
                    ]
                  }
                ]
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Test Agent',
        agentInstructions: 'You are a helpful test agent',
        welcomeMessage: 'Hello! I am your test agent.',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test/documents/folder',
        description: 'A test agent'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data,
      {
        schemaVersion: "0.2.0",
        customCopilotConfig: {
          conversationStarters: {
            conversationStarterList: [],
            welcomeMessage: {
              text: "Hello! I am your test agent."
            }
          },
          gptDefinition: {
            name: "Test Agent",
            description: "A test agent",
            instructions: "You are a helpful test agent",
            capabilities: [
              {
                name: "OneDriveAndSharePoint",
                items_by_sharepoint_ids: [
                ],
                items_by_url: [
                  {
                    url: "https://contoso.sharepoint.com/sites/test/documents/folder",
                    name: "Test Folder",
                    site_id: "f1e2d3c4-b5a6-7890-1234-56789abcdef0",
                    web_id: "123e4567-e89b-12d3-a456-426614174000",
                    list_id: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a",
                    unique_id: "0f1e2d3c-4b5a-6789-0123-456789abcdef",
                    type: "Folder"
                  }
                ]
              }
            ]
          },
          icon: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAYAAACOEfKtAAAIJklEQVR4nO2bCWwUVRjHv9m73RYEBEQR5RLKpSgKqHggtBQRBAQFEY94GzUcVlBDQIkhGhUFJCqCBFDuUzkrN4qCii1GG4UKqFwCgba7OzO7M37fN3a729mFbh+IJO+XkG135r33vf981xtAOTGgnwmSauM43wZc6EgBBZECCiIFFEQKKIgUUBApoCBSQEGkgIJIAQWRAgoiBRRECiiIFFAQKaAgUkBBpICCSAEFkQIKIgUURAooiBRQECmgIFJAQaSAgkgBBZECCiIFFMQuoOMC0NTlAsXjOd9WMK7KX6Q/+TQYhw+B+vkKMFX1fNiUGEUBb7ds8OTmgvOyhvy7cewY6Fs3Q2jRQjBDoaTj3FdfA+7ON4Kz0RUATieU5I2w3eaoXx8yXnol6fL6zp0QnDXT9r1NQMXnA9+9g8DTPQdC8+eCtnEDgGGksNNzQ/rjT4Lnjm5glpSAtmUz7kgHV5u24O3TF5wtsqB03BiASCRujJKRAf7nh4ELBSTMU6cggs6RCEedi8HR4FJrrwn2q9SokXCcTUBt+9fgbJkFjtq12Ru9Pe+E0OxZoO/6IeVNny1cWa1YPOPwYSh5eTQKcdK64HZD5rjXwNWyJbg7dgL9q20Vg9DT/C++BK4WLSBcWADBObMgsndv0jUctWrxZ2jJIgjNm1tl22wJj4woee4Z9L55YAaD7PZ+dO2MV8aA88rGVZ7Ymt0Rl1MddeuBs1nzqLGx9zkbXg7Opk1B8WfYpnG1b8+fav7aCvHYWB20zZuse1q0jBvjRcFJPP3776D09fGnFY9Q/rWJ0kIq2DyQoHwSWjgf1HVrwNd/AOae7uBqdzVkTmjL4ROa+xku9PcZJ6ecQuMCUyaDN7cnOJs0iV6jjQUmvcvX04Y+hCFUx7qA4aNtWA+BaR9hSIb5K3XFctDy8zF8T9ltRREZ9LhYPLgehXSQ54nYxlXGUau2NV+KAjpHtc4am/QqFpHwD9+DtnUrKDVrojc2Yi/0ZGeD4vXhU93DXpAMzy23YnK+BNwdOkBk/z7QUZhI8V5wXtIAnI0bgxs9y5uTC+GiX0DfiNcOHEBPbGh5E84b/uVnayJNA7OsDCAcjl8AC0Ta4Ae4AKhYSIy//rLEqFsX0u4bzOPDu3eDD/OkN6cHuK+9jqs3rWOz9fauvD8a4+lyC3i73gGutu1AwQdqHEqcN9mEVP6NNHkQGexq145/p6TMnpq/zr45hMKePExdtRKCMz6umOfyRpD51jv8s7p6FQSnT4tec7e/FvyjX2bBS0YOt9vQHFNA7Toc6p4uXcDVug1omzawl0fn6HA9+PNGgXHkiOXZKDR1FFQg6Wd9x7dQ9tabccUiY+yr4GrVOuG+KeoCk9/DDdulShjCyaA8Ujp+nBV29w9BL0JBH3mUwzP46RzQv9mecBwZHDfPgf1gnDjOYUMCxlLudZQvE+Hr1ZtbklibgjM/ibuHKip/1qvHOTKEBcQ4cYKjxz8yD9zX3wDe7tmgrlkdHUO5zwwEQF2+FDQqRhj25LG+IUPZIyNFRaCuXQ2VqVbXHC74EUpG5UHgvYmWEFj+/SNe4GpZ9UksjzUDZfHfU76iJ52koQ/hBsvemACB9yfzA6OoyJzwBqaYiypu8nqtJTB8A1MmsXg89e/FmFs/4J89XbvFzUv5+OTDQyG0eBGHrHH0KAsc/NiKDk92TkJ7qn3sUGpgTsRc5ci0+iNTxzwVDKQ+UYr/ySKyZw82tTu4P6UwVJcvQ0+rD75+/WOMsz4ot1YOu3BhIYcz5Ts60cTbYjdG27YFB+mcmxOdflIKYbYtLQ28GEbeu3pbOcU0/q3Mn/JTSxkl9SGxUIh6e/eJ834KRZ46Pd0+AD2c2jMqigoKaCbI3XFgMTMDQauRJgGxoMVSdQFxMWpnqK2hxYnw7kIIzp5lVePqUgUPTHvsCT6OBT6Yyk1x/Hj7BOU9n6v5VbZrJKojM5OrevlRlXPd0AfZq9WlS2z3K34/mJrKwlfmzALSWfLGm7AtGMQtCRuIRSA0Zzb3cv8FJuYwKgjUFlUW0N2xo2XTb79Gv6MHGvnzD2zam3FV17EVK8eD52nqGfUfd0XF5/YJ87g3uwdomPdihaL2hu4PFxQk7CdPKyD1QVxtmzTl3ykZ0wlFw56tKs3p2UJdu4aTuOfW2zgSuNpj6PFZuEcubziEuTAKChOcMZ0b+fRhI9CrFmNbtB+PfFmYfnqx56kLF0RvN44eAW39l3xcpHZGXfkFn7mdmBZ8ve7ivaqLF9gNgyQCcnuCwlG7wvbgyURdsYxPBEnfepxD6PhWOnYMns2fAs9NN/Ofcsh7AlOnYOU8GDeGOoWyiW9DOoa/D5vqckgs6hnJQ2MJYLU1sVjQG5/0Z56tuB+dJjjtQyxIRQltszXSvnsGgm/AQA5dMCL4ZNZDaMG8aCuQCtzEYktBx6PKr8bo9EChYRw6zOtUWKRgW9SAvcg4eBAqQ9fo3ExtDr12i+zblzAPRqfD9alboFxGqSBMoX6a6KH87mrajF9UmCdPWvefptDYBPQPHwnuTp1B/24n57nIH/Zjj6QCWwhHiou5gQz/tPt82HPBYROQ3odJqk7KjXRVobfBvrv74gru6k+CuZH+asE4fvzsGXaWOacC8qt0QQG1bXiw/x8LmNLrLImdC+DvMP/fSAEFkQIKIgUURAooiBRQECmgIFJAQaSAgkgBBZECCiIFFEQKKIgUUBApoCBSQEGkgIJIAQWRAgoiBRRECiiIFFAQKaAgUkBBpICCSAEFkQIKIgUU5B91HS13TtrWPgAAAABJRU5ErkJggg=="
        }
      }
    );
  });

  it('handles API errors properly', async () => {
    const errorMessage = 'Agent creation failed';

    sinon.stub(request, 'post').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Error Agent',
        agentInstructions: 'This will fail',
        welcomeMessage: 'This should not work',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test'
      }
    }), new CommandError(errorMessage));
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'invalid-url',
      name: 'Test Agent',
      agentInstructions: 'Test instructions',
      welcomeMessage: 'Test welcome',
      sourceUrls: 'https://contoso.sharepoint.com',
      description: 'A test agent'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if name is empty', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      agentInstructions: 'Test instructions',
      welcomeMessage: 'Test welcome',
      sourceUrls: 'https://contoso.sharepoint.com',
      description: 'A test agent'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if description is empty', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      name: 'Test Agent',
      agentInstructions: 'Test instructions',
      welcomeMessage: 'Test welcome',
      sourceUrls: 'https://contoso.sharepoint.com'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if agentInstructions is empty', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      name: 'Test Agent',
      welcomeMessage: 'Test welcome',
      sourceUrls: 'https://contoso.sharepoint.com',
      description: 'A test agent'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if welcomeMessage is empty', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      name: 'Test Agent',
      agentInstructions: 'Test instructions',
      sourceUrls: 'https://contoso.sharepoint.com',
      description: 'A test agent'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if sourceUrls is empty', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      name: 'Test Agent',
      agentInstructions: 'Test instructions',
      welcomeMessage: 'Test welcome',
      description: 'A test agent'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if sourceUrls contains invalid SharePoint URLs', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      name: 'Test Agent',
      agentInstructions: 'Test instructions',
      welcomeMessage: 'Test welcome',
      sourceUrls: 'https://contoso.sharepoint.com,invalid-url,https://contoso.sharepoint.com/sites/docs',
      description: 'A test agent'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with all required options', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      name: 'Test Agent',
      agentInstructions: 'Test instructions',
      welcomeMessage: 'Test welcome',
      sourceUrls: 'https://contoso.sharepoint.com',
      description: 'A test agent'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with multiple valid source URLs', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      name: 'Test Agent',
      agentInstructions: 'Test instructions',
      welcomeMessage: 'Test welcome',
      sourceUrls: 'https://contoso.sharepoint.com/sites/test,https://contoso.sharepoint.com/sites/docs,https://contoso.sharepoint.com/sites/team',
      description: 'A test agent'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with all options including optional ones', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      name: 'Complete Agent',
      agentInstructions: 'Comprehensive test instructions',
      welcomeMessage: 'Welcome to the complete test',
      sourceUrls: 'https://contoso.sharepoint.com/sites/test,https://contoso.sharepoint.com/sites/docs',
      description: 'A complete test agent',
      icon: 'https://contoso.sharepoint.com/sites/test/SiteAssets/icon.png',
      conversationStarters: 'Hello,How can I help?,What do you need?'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with valid icon URL', async () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com/sites/test',
      name: 'Test Agent',
      agentInstructions: 'Test instructions',
      welcomeMessage: 'Test welcome',
      sourceUrls: 'https://contoso.sharepoint.com',
      icon: 'https://example.com/icon.png',
      description: 'A test agent'
    });
    assert.strictEqual(actual.success, true);
  });

  it('handles empty UniqueID correctly', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='/sites/test/SiteAssets/Copilots')/Files/AddUsingPath(DecodedUrl='Test%20Agent.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)"
      ) {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/web/lists/EnsureSiteAssetsLibrary()') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/search/postquery') {
        return {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: [
                  {
                    Cells: [
                      { Key: "contentclass", Value: "STS_ListItem_DocumentLibrary" },
                      { Key: "Title", Value: "Test Document" },
                      { Key: "Path", Value: "https://contoso.sharepoint.com/sites/test/Shared Documents/Test Document.docx" },
                      { Key: "SiteName", Value: "Test Site" },
                      { Key: "SiteTitle", Value: "Test Site" },
                      { Key: "ListID", Value: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a" },
                      { Key: "ListItemID", Value: "a7c6b5d4-e3f2-1a09-b8c7-6e5d4c3b2a1f" },
                      { Key: "SiteID", Value: "f1e2d3c4-b5a6-7890-1234-56789abcdef0" },
                      { Key: "WebId", Value: "123e4567-e89b-12d3-a456-426614174000" },
                      { Key: "IsDocument", Value: "true" },
                      { Key: "IsContainer", Value: "false" }
                    ]
                  }
                ]
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Test Agent',
        agentInstructions: 'You are a helpful test agent',
        welcomeMessage: 'Hello! I am your test agent.',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test',
        description: 'A test agent'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data.customCopilotConfig.gptDefinition.capabilities[0].items_by_sharepoint_ids[0].unique_id, "");
  });

  it('handles UniqueID without curly braces correctly', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/sites/test/_api/web/GetFolderByServerRelativePath(DecodedUrl='/sites/test/SiteAssets/Copilots')/Files/AddUsingPath(DecodedUrl='Test%20Agent.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)"
      ) {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/web/lists/EnsureSiteAssetsLibrary()') {
        return;
      }

      if (opts.url === 'https://contoso.sharepoint.com/sites/test/_api/search/postquery') {
        return {
          PrimaryQueryResult: {
            RelevantResults: {
              Table: {
                Rows: [
                  {
                    Cells: [
                      { Key: "contentclass", Value: "STS_ListItem_DocumentLibrary" },
                      { Key: "Title", Value: "Test Document" },
                      { Key: "Path", Value: "https://contoso.sharepoint.com/sites/test/Shared Documents/Test Document.docx" },
                      { Key: "SiteName", Value: "Test Site" },
                      { Key: "SiteTitle", Value: "Test Site" },
                      { Key: "ListID", Value: "b1a5e7c2-3d4f-4e6a-9b8c-2f3e4d5c6b7a" },
                      { Key: "ListItemID", Value: "a7c6b5d4-e3f2-1a09-b8c7-6e5d4c3b2a1f" },
                      { Key: "SiteID", Value: "f1e2d3c4-b5a6-7890-1234-56789abcdef0" },
                      { Key: "WebId", Value: "123e4567-e89b-12d3-a456-426614174000" },
                      { Key: "UniqueID", Value: "0f1e2d3c-4b5a-6789-0123-456789abcdef" },
                      { Key: "IsDocument", Value: "true" },
                      { Key: "IsContainer", Value: "false" }
                    ]
                  }
                ]
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/test',
        name: 'Test Agent',
        agentInstructions: 'You are a helpful test agent',
        welcomeMessage: 'Hello! I am your test agent.',
        sourceUrls: 'https://contoso.sharepoint.com/sites/test',
        description: 'A test agent'
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data.customCopilotConfig.gptDefinition.capabilities[0].items_by_sharepoint_ids[0].unique_id, "0f1e2d3c-4b5a-6789-0123-456789abcdef");
  });
});
