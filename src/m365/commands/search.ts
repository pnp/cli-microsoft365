import { SearchHit, SearchResponse } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../cli/Logger.js';
import { globalOptionsZod } from '../../Command.js';
import request, { CliRequestOptions } from '../../request.js';
import { ODataResponse } from '../../utils/odata.js';
import GraphCommand from '../base/GraphCommand.js';
import commands from './commands.js';

const allowedScopes = ['chatMessage', 'message', 'event', 'drive', 'driveItem', 'list', 'listItem', 'site', 'bookmark', 'acronym', 'person'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  queryText: z.string().optional().alias('q'),
  scopes: z.string()
    .refine(value => value.split(',').map(x => x.trim()).every(scope => (allowedScopes as readonly string[]).includes(scope)), {
      error: e => {
        const scopes = (e.input as string).split(',').map(x => x.trim());
        const invalidScope = scopes.find(scope => !(allowedScopes as readonly string[]).includes(scope));
        return `'${invalidScope}' is not a valid scope. Allowed scopes are ${allowedScopes.join(', ')}.`;
      }
    }).alias('s'),
  startIndex: z.number()
    .refine(n => n >= 0, {
      error: e => `'${e.input}' is not a valid value for option 'startIndex'. Start index must be greater or equal to 0.`
    }).optional(),
  pageSize: z.number()
    .refine(n => n >= 1 && n <= 500, {
      error: e => `'${e.input}' is not a valid value for option 'pageSize'. Page size must be between 1 and 500.`
    }).optional(),
  allResults: z.boolean().optional(),
  resultsOnly: z.boolean().optional(),
  enableTopResults: z.boolean().optional(),
  select: z.string().optional(),
  sortBy: z.string().optional(),
  enableSpellingSuggestion: z.boolean().optional(),
  enableSpellingModification: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SearchCommand extends GraphCommand {
  public get name(): string {
    return commands.SEARCH;
  }

  public get description(): string {
    return 'Uses the Microsoft Search to query Microsoft 365 data';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => {
        if (opts.sortBy) {
          const scopes = opts.scopes.split(',').map(x => x.trim());
          return !scopes.some(scope => scope === 'message' || scope === 'event');
        }
        return true;
      }, {
        error: 'Sorting the results is not supported for messages and events.'
      })
      .refine(opts => {
        if (opts.enableTopResults) {
          const scopes = opts.scopes.split(',').map(x => x.trim());
          if (scopes.length === 1) {
            return scopes[0] === 'message' || scopes[0] === 'chatMessage';
          }
          if (scopes.length === 2) {
            return scopes.includes('message') && scopes.includes('chatMessage');
          }
          return false;
        }
        return true;
      }, {
        error: 'Top results are only supported for messages and chat messages.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let result: SearchResponse;
    const searchHits: SearchHit[] = [];

    try {
      let allResults = args.options.allResults ? args.options.allResults : false;
      let startIndex = args.options.startIndex ? args.options.startIndex : 0;
      const pageSize = args.options.pageSize ? args.options.pageSize : 25;

      do {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/search/query`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            requests: [
              {
                "entityTypes": args.options.scopes.split(',').map(scope => scope.trim()),
                "query": {
                  "queryString": args.options.queryText ?? '*'
                },
                "enableTopResults": args.options.enableTopResults,
                "from": startIndex,
                "size": pageSize,
                "fields": this.getProperties(args.options),
                "sortProperties": this.getSortProperties(args.options),
                "queryAlterationOptions": {
                  "enableModification": args.options.enableSpellingModification,
                  "enableSuggestion": args.options.enableSpellingSuggestion
                }
              }
            ]
          }
        };

        const response = await request.post<ODataResponse<SearchResponse>>(requestOptions);
        result = response.value[0];

        if (allResults && result.hitsContainers) {
          allResults = result.hitsContainers[0].moreResultsAvailable!;
        }

        if (allResults) {
          startIndex += pageSize;
        }

        if (result.hitsContainers && result.hitsContainers[0].hits) {
          searchHits.push(...result.hitsContainers[0].hits);
        }
      }
      while (allResults);

      if (args.options.resultsOnly) {
        await logger.log(searchHits);
      }
      else {
        if (result.hitsContainers && result.hitsContainers[0].hits) {
          result.hitsContainers[0].hits = searchHits;
        }

        await logger.log(result);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getProperties(options: Options): string[] | undefined {
    if (!options.select) {
      return undefined;
    }

    return options.select.split(',').map(prop => prop.trim());
  }

  private getSortProperties(options: Options): any[] | undefined {
    if (!options.sortBy) {
      return undefined;
    }

    const properties = options.sortBy.split(',').map(prop => prop.trim()).map(property => {
      const sortDefinitions = property.split(':');
      const name = sortDefinitions[0];
      let isDescending = false;

      if (sortDefinitions.length === 2) {
        const order = sortDefinitions[1].trim();
        isDescending = order === 'desc';
      }

      return {
        "name": name,
        "isDescending": isDescending
      };
    });

    return properties;
  }
}

export default new SearchCommand();