/* eslint-disable camelcase */
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { spo } from '../../../../utils/spo.js';
import { formatting } from '../../../../utils/formatting.js';
import { SearchResult } from '../search/datatypes/SearchResult.js';
import { ResultTableRow } from '../search/datatypes/ResultTableRow.js';
import { SearchResultProperty } from '../search/datatypes/SearchResultProperty.js';

const options = globalOptionsZod
  .extend({
    webUrl: zod.alias('u', z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
      message: 'Specify a valid SharePoint site URL'
    })),
    name: zod.alias('n', z.string()),
    agentInstructions: zod.alias('a', z.string()),
    welcomeMessage: zod.alias('w', z.string()),
    sourceUrls: zod.alias('s', z.string().refine(urls => {
      const urlArray = urls.split(',').map(url => url.trim());
      return urlArray.every(url => url && validation.isValidSharePointUrl(url) === true);
    }, {
      message: 'All source URLs must be valid SharePoint URLs'
    })),
    description: zod.alias('d', z.string()),
    icon: zod.alias('i', z.string().optional()),
    conversationStarters: zod.alias('c', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface AgentSource {
  url: string;
  name: string;
  site_id: string;
  web_id: string;
  list_id: string;
  unique_id: string;
  type: string;
}

const urlSourceMap: Record<string, string> = {
  STS_ListItem_DocumentLibrary: 'File',
  STS_List_DocumentLibrary: 'List',
  STS_Web: 'Site',
  STS_Site: 'Site'
};

interface AgentCapabilities {
  name: string;
  items_by_sharepoint_ids: AgentSource[];
  items_by_url: AgentSource[];
}

interface AgentRequestBody {
  schemaVersion: string;
  customCopilotConfig: {
    conversationStarters: {
      conversationStarterList: Array<{ text: string }>;
      welcomeMessage: { text: string };
    };
    gptDefinition: {
      name: string;
      description: string;
      instructions: string;
      capabilities: AgentCapabilities[];
    };
    icon: string;
  };
}

class SpoAgentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.AGENT_ADD;
  }

  public get description(): string {
    return 'Adds a new SharePoint agent';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Adding SharePoint agent '${args.options.name}' to site '${args.options.webUrl}'...`);
      }

      await this.ensureSiteAssetsLibrary(args.options.webUrl, logger);
      await spo.ensureFolder(args.options.webUrl, 'SiteAssets/Copilots', logger, this.verbose);

      const sourceUrls = args.options.sourceUrls.split(',');

      const capabilities: AgentCapabilities = await this.resolveSourceUrls(sourceUrls, args.options.webUrl, logger);

      const conversationStartersArray = args.options.conversationStarters
        ? args.options.conversationStarters.split(',').map((starter: string) => { return { text: starter }; })
        : [];

      const cliM365DefaultIcon = 'data:image/png;base64,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';

      const requestBody: AgentRequestBody = {
        schemaVersion: "0.2.0",
        customCopilotConfig: {
          conversationStarters: {
            conversationStarterList: conversationStartersArray,
            welcomeMessage: {
              text: args.options.welcomeMessage
            }
          },
          gptDefinition:
          {
            name: args.options.name,
            description: args.options.description,
            instructions: args.options.agentInstructions,
            capabilities: [capabilities]
          },
          icon: args.options.icon || cliM365DefaultIcon
        }
      };

      const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, '/SiteAssets/Copilots/');

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${serverRelativePath}')/Files/AddUsingPath(DecodedUrl='${formatting.encodeQueryParameter(args.options.name)}.agent',EnsureUniqueFileName=true,AutoCheckoutOnInvalidData=true)`,
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const result = await request.post(requestOptions);

      if (this.verbose) {
        await logger.logToStderr(`Agent '${args.options.name}' has been successfully created.`);
      }

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async ensureSiteAssetsLibrary(webUrl: string, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Ensuring Site Assets library exists at ${webUrl}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/lists/EnsureSiteAssetsLibrary()`,
      headers: {
        'Accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private async resolveSourceUrls(sourceUrls: string[], webUrl: string, logger: Logger): Promise<AgentCapabilities> {
    const resolvedUrls: AgentSource[] = [];
    const resolvedFiles: AgentSource[] = [];

    for (const sourceUrl of sourceUrls) {
      if (this.verbose) {
        await logger.logToStderr(`Resolving source URL: ${sourceUrl}`);
      }

      const requestBody = {
        request: {
          QueryTemplate: "({searchterms}) (contentclass:STS_Web OR contentclass:STS_Site OR contentclass:STS_ListItem_DocumentLibrary OR contentclass:STS_List_DocumentLibrary)",
          Querytext: `Path=\"${sourceUrl}\"`,
          SelectProperties: ["contentclass", "Title", "Path", "SiteName", "SiteTitle", "ListID", "ListItemID", "SiteID", "WebId", "UniqueID", "IsDocument", "IsContainer"],
          RowLimit: 1,
          TrimDuplicates: false
        }
      };

      const requestOptions: CliRequestOptions = {
        url: `${webUrl}/_api/search/postquery`,
        headers: {
          'Accept': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: requestBody
      };

      const response: SearchResult = await request.post(requestOptions);

      if (response.PrimaryQueryResult.RelevantResults.Table.Rows.length === 0) {
        await logger.logToStderr(`${sourceUrl} has been skipped because no results were found.`);
        continue;
      }

      const row = response.PrimaryQueryResult.RelevantResults.Table.Rows[0];
      const isContainer = this.getCellValue(row, "IsContainer");

      let uniqueId = this.getCellValue(row, "UniqueID");
      if (uniqueId.startsWith('{') && uniqueId.endsWith('}')) {
        uniqueId = uniqueId.slice(1, -1);
      }

      const contentClass = this.getCellValue(row, "contentclass");

      const resolvedItem = {
        url: sourceUrl,
        name: this.getCellValue(row, "Title"),
        site_id: this.getCellValue(row, "SiteID"),
        web_id: this.getCellValue(row, "WebId"),
        list_id: this.getCellValue(row, "ListID") || '',
        unique_id: uniqueId,
        type: isContainer === 'true' && contentClass === 'STS_ListItem_DocumentLibrary' ? 'Folder' : urlSourceMap[contentClass]
      };

      if (isContainer === 'false' && contentClass === 'STS_ListItem_DocumentLibrary') {
        resolvedFiles.push(resolvedItem);
      }
      else {
        resolvedUrls.push(resolvedItem);
      }
    }

    return {
      name: "OneDriveAndSharePoint",
      items_by_sharepoint_ids: resolvedFiles,
      items_by_url: resolvedUrls
    };
  }

  private getCellValue(row: ResultTableRow, key: string): string {
    return row.Cells.find((cell: SearchResultProperty) => cell.Key === key)?.Value || '';
  }
}

export default new SpoAgentAddCommand();
