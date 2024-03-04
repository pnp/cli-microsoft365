import { SearchHit, SearchResponse } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../request.js';
import GraphCommand from '../../base/GraphCommand.js';
import GlobalOptions from '../../../GlobalOptions.js';
import { ODataResponse } from '../../../utils/odata.js';

const commandName = 'search';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  queryString?: string;
  scopes: string;
  startIndex?: number;
  pageSize?: number;
  allResults?: boolean;
  resultsOnly?: boolean;
  enableTopResults?: boolean;
  select?: string;
  sortBy?: string;
  enableSpellingSuggestion?: boolean;
  enableSpellingModification?: boolean;
}

class SearchSearchCommand extends GraphCommand {
  private allowedScopes: string[] = ['chatMessage', 'message', 'event', 'drive', 'driveItem', 'list', 'listItem', 'site', 'bookmark', 'acronym', 'person'];

  public get name(): string {
    return commandName;
  }

  public get description(): string {
    return 'Uses the Microsoft Search to query Microsoft 365 data';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        query: typeof args.options.query !== 'undefined',
        startIndex: typeof args.options.startIndex !== 'undefined',
        pageSize: typeof args.options.pageSize !== 'undefined',
        allResults: !!args.options.allResults,
        resultsOnly: !!args.options.resultsOnly,
        enableTopResults: !!args.options.enableTopResults,
        select: typeof args.options.select !== 'undefined',
        sortBy: typeof args.options.sortBy !== 'undefined',
        enableSpellingSuggestion: !!args.options.enableSpellingSuggestion,
        enableSpellingModification: !!args.options.enableSpellingModification
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-q --queryString [queryString]'
      },
      {
        option: '-s, --scopes <scopes>',
        autocomplete: this.allowedScopes
      },
      {
        option: '--startIndex [startIndex]'
      },
      {
        option: '--pageSize [pageSize]'
      },
      {
        option: '--allResults'
      },
      {
        option: '--resultsOnly'
      },
      {
        option: '--enableTopResults'
      },
      {
        option: '--select [select]'
      },
      {
        option: '--sortBy [sortBy]'
      },
      {
        option: '--enableSpellingSuggestion'
      },
      {
        option: '--enableSpellingModification'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const scopes = args.options.scopes.split(',').map(x => x.trim());

        if (!scopes.every(scope => this.allowedScopes.indexOf(scope) > -1)) {
          const invalidScope = scopes.find(scope => this.allowedScopes.indexOf(scope) === -1);
          return `'${invalidScope}'' is not a valid scope. Allowed scopes are ${this.allowedScopes.join(', ')}.`;
        }

        if (args.options.startIndex !== undefined && args.options.startIndex < 0) {
          return `'${args.options.startIndex}' is not a valid value for option 'startIndex'. Start index must be greater or equal to 0.`;
        }

        if (args.options.pageSize !== undefined && (args.options.pageSize < 1 || args.options.pageSize > 500)) {
          return `'${args.options.pageSize}' is not a valid value for option 'pageSize'. Page size must be between 1 and 500.`;
        }

        if (args.options.sortBy && scopes.some(scope => scope === 'message' || scope === 'event')){
          return 'Sorting the results is not supported for messages and events.';
        }

        if (args.options.enableTopResults &&
          ((scopes.length === 1 && scopes.indexOf('message') === -1 && scopes.indexOf('chatMessage') === -1) ||
          (scopes.length === 2) && !(scopes.indexOf('message') > -1 && scopes.indexOf('chatMessage') > -1))) {
          return 'Top results are only supported for messages and chat messages.';
        }

        return true;
      }
    );
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
                  "queryString": args.options.queryString ?? '*'
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

export default new SearchSearchCommand();