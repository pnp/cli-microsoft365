import { ExternalConnectors, SearchResponse } from '@microsoft/microsoft-graph-types';
import os from 'os';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { settingsNames } from '../../../../settingsNames.js';
import { ODataResponse } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  ux?: string;
}

/**
 * Defines a check that can be run by the doctor command
 */
interface Check {
  // id to store the data to pass between checks
  id: string;
  // text to display to the user
  text: string;
  // function to run; only for automated checks
  fn?: (id: string, args: CommandArgs) => Promise<CheckResult>;
  type: 'required' | 'recommended'
}

/**
 * Defines the result of a check
 */
interface CheckResult {
  id: string;
  // data that the check wants to pass to other checks
  data?: any;
  // error object that occurred while running the check
  error?: any;
  // error message to display to the user
  errorMessage?: string;
  // if true, the doctor command will stop running checks
  shouldStop?: boolean;
  status: 'passed' | 'failed' | 'manual';
}

class ExternalConnectionDoctorCommand extends GraphCommand {
  private checksStatus: CheckResult[] = [];
  private static readonly supportedUx: string[] = ['copilot', 'search', 'all'];

  public get name(): string {
    return commands.CONNECTION_DOCTOR;
  }

  public get description(): string {
    return 'Checks if the external connection is correctly configured for use with the specified Microsoft 365 experience';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--ux [ux]',
        autocomplete: ExternalConnectionDoctorCommand.supportedUx
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.ux) {
          if (!ExternalConnectionDoctorCommand.supportedUx.find(u => u === args.options.ux)) {
            return `${args.options.ux} is not a valid UX. Allowed values are ${ExternalConnectionDoctorCommand.supportedUx.join(', ')}`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const ux = args.options.ux ?? 'all';
    const output = args.options.output;
    this.checksStatus = [];

    const showSpinner = cli.getSettingWithDefaultValue<boolean>(settingsNames.showSpinner, true) &&
      output === 'text' &&
      typeof global.it === 'undefined';

    let checks: Check[] = [
      {
        id: 'loadExternalConnection',
        text: 'Load connection',
        fn: this.loadConnection,
        type: 'required'
      },
      {
        id: 'loadSchema',
        text: 'Load schema',
        fn: this.loadSchema,
        type: 'required'
      }
    ];

    if (ux === 'copilot' || ux === 'all') {
      checks.push(
        {
          id: 'copilotRequiredSemanticLabels',
          text: 'Required semantic labels',
          fn: this.checkCopilotRequiredSemanticLabels,
          type: 'required'
        },
        {
          id: 'searchableProperties',
          text: 'Searchable properties',
          fn: this.checkSearchableProperties,
          type: 'required'
        },
        {
          id: 'contentIngested',
          text: 'Items have content ingested',
          fn: this.checkContentIngested,
          type: 'required'
        },
        {
          id: 'enabledForInlineResults',
          text: 'Connection configured for inline results',
          type: 'required'
        },
        {
          id: 'itemsHaveActivities',
          text: 'Items have activities recorded',
          type: 'recommended'
        },
        {
          id: 'meaningfulNameAndDescription',
          text: 'Meaningful connection name and description',
          type: 'required'
        }
      );
    }

    if (ux === 'search' || ux === 'all') {
      checks.push(
        {
          id: 'semanticLabels',
          text: 'Semantic labels',
          fn: this.checkSemanticLabels,
          type: 'recommended'
        },
        {
          id: 'searchableProperties',
          text: 'Searchable properties',
          fn: this.checkSearchableProperties,
          type: 'recommended'
        },
        {
          id: 'resultType',
          text: 'Result type',
          fn: this.checkResultType,
          type: 'recommended'
        },
        {
          id: 'contentIngested',
          text: 'Items have content ingested',
          fn: this.checkContentIngested,
          type: 'recommended'
        },
        {
          id: 'itemsHaveActivities',
          text: 'Items have activities recorded',
          type: 'recommended'
        }
      );
    }

    checks.push(
      {
        id: 'urlToItemResolver',
        text: 'urlToItemResolver configured',
        fn: this.checkUrlToItemResolverConfigured,
        type: 'recommended'
      }
    );

    // filter out duplicate checks based on their IDs
    checks = checks.filter((check, index, self) => self.findIndex(c => c.id === check.id) === index);

    for (const check of checks) {
      if (this.debug) {
        logger.logToStderr(`Running check ${check.id}...`);
      }

      // don't show spinner if running tests
      /* c8 ignore next 3 */
      if (showSpinner) {
        cli.spinner.start(check.text);
      }

      // only automated checks have functions
      if (!check.fn) {
        // don't show spinner if running tests
        /* c8 ignore next 3 */
        if (showSpinner) {
          cli.spinner.info(`${check.text} (manual)`);
        }

        this.checksStatus.push({
          ...check,
          status: 'manual'
        });

        continue;
      }

      const result = await check.fn.bind(this)(check.id, args);
      this.checksStatus.push({ ...check, ...result });

      if (result.status === 'passed') {
        // don't show spinner if running tests
        /* c8 ignore next 3 */
        if (showSpinner) {
          cli.spinner.succeed();
        }

        continue;
      }

      if (result.status === 'failed') {
        // don't show spinner if running tests
        /* c8 ignore next 9 */
        if (showSpinner) {
          const message = `${check.text}: ${result.errorMessage}`;
          if (check.type === 'required') {
            cli.spinner.fail(message);
          }
          else {
            cli.spinner.warn(message);
          }
        }

        if (result.shouldStop) {
          break;
        }
      }
    }

    if (output === 'text' || output === 'none') {
      return;
    }

    this.checksStatus.forEach(s => {
      delete s.data;
      delete (s as any).fn;
      delete s.shouldStop;
    });

    if (output === 'json' || output === 'md') {
      await logger.log(this.checksStatus);
      return;
    }

    if (output === 'csv') {
      this.checksStatus.forEach(r => {
        // we need to set errorMessage to empty string so that it's not
        // removed from the CSV output
        r.errorMessage = r.errorMessage ?? '';
      });
      await logger.log(this.checksStatus);
    }
  }

  public getMdOutput(logStatement: any[], command: Command, options: GlobalOptions): string {
    const output: string[] = [
      `# ${command.getCommandName()} ${Object.keys(options).filter(o => o !== 'output').map(k => `--${k} "${options[k]}"`).join(' ')}`, os.EOL,
      os.EOL,
      `Date: ${(new Date().toLocaleDateString())}`, os.EOL,
      os.EOL
    ];

    if (logStatement && logStatement.length > 0) {
      const properties = ['text', 'type', 'status', 'errorMessage'];

      output.push('Check|Type|Status|Error message', os.EOL);
      output.push(':----|:--:|:----:|:------------', os.EOL);
      logStatement.forEach(r => {
        output.push(properties.map(p => r[p] ?? '').join('|'), os.EOL);
      });
      logStatement.push(os.EOL);
    }

    return output.join('').trimEnd();
  }

  private async loadConnection(id: string, args: CommandArgs): Promise<CheckResult> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/external/connections/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const externalConnection = await request.get<ExternalConnectors.ExternalConnection>(requestOptions);
      return {
        id,
        data: externalConnection,
        status: 'passed'
      };
    }
    catch (ex) {
      return {
        id,
        error: (ex as any)?.response?.data?.error?.innerError?.message,
        errorMessage: 'Connection not found',
        shouldStop: true,
        status: 'failed'
      };
    }
  }

  private async loadSchema(id: string, args: CommandArgs): Promise<CheckResult> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/external/connections/${args.options.id}/schema`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        prefer: 'include-unknown-enum-members'
      },
      responseType: 'json'
    };

    try {
      const schema = await request.get<ExternalConnectors.Schema>(requestOptions);
      return {
        id,
        data: schema,
        status: 'passed'
      };
    }
    catch (ex) {
      return {
        id,
        errorMessage: 'Schema not found',
        error: (ex as any)?.response?.data?.error?.innerError?.message,
        shouldStop: true,
        status: 'failed'
      };
    }
  }

  private async checkCopilotRequiredSemanticLabels(id: string): Promise<CheckResult> {
    const schema: ExternalConnectors.Schema = this.checksStatus.find(r => r.id === 'loadSchema')!.data;
    const requiredLabels: ExternalConnectors.Label[] = ['title', 'url', 'iconUrl'];

    for (const label of requiredLabels) {
      if (!schema.properties?.find(p => p.labels?.find(l => l.toString() === label))) {
        return {
          id,
          errorMessage: `Missing label ${label}`,
          status: 'failed'
        };
      }
    }

    return {
      id,
      status: 'passed'
    };
  }

  private async checkSearchableProperties(id: string): Promise<CheckResult> {
    const schema: ExternalConnectors.Schema = this.checksStatus.find(r => r.id === 'loadSchema')!.data;

    if (!schema.properties?.some(p => p.isSearchable)) {
      return {
        id,
        errorMessage: 'Schema does not have any searchable properties',
        status: 'failed'
      };
    }

    return {
      id,
      status: 'passed'
    };
  }

  private async checkContentIngested(id: string, args: CommandArgs): Promise<CheckResult> {
    try {
      // find items that belong to the connection
      const searchRequestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/search/query`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          requests: [
            {
              entityTypes: [
                'externalItem'
              ],
              contentSources: [
                `/external/connections/${args.options.id}`
              ],
              query: {
                queryString: '*'
              },
              from: 0,
              size: 1
            }
          ]
        }
      };
      const result = await request.post<ODataResponse<SearchResponse>>(searchRequestOptions);

      const hit = result.value?.[0].hitsContainers?.[0]?.hits?.[0];
      if (!hit) {
        return {
          id,
          errorMessage: 'No items found that belong to the connection',
          status: 'failed'
        };
      }

      // something@tenant,itemId
      const itemId = (hit.resource as any)?.properties?.substrateContentDomainId?.split(',')?.[1];
      if (!itemId) {
        return {
          id,
          errorMessage: 'Item does not have substrateContentDomainId property or the property is invalid',
          status: 'failed'
        };
      }

      const externalItemRequestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/external/connections/${args.options.id}/items/${itemId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const externalItem = await request.get<ExternalConnectors.ExternalItem>(externalItemRequestOptions);
      if (!externalItem.content?.value) {
        return {
          id,
          data: externalItem,
          errorMessage: 'Item does not have content or content is empty',
          status: 'failed'
        };
      }

      return {
        id,
        data: externalItem,
        status: 'passed'
      };
    }
    catch (ex) {
      return {
        id,
        error: (ex as any)?.response?.data?.error?.innerError?.message,
        errorMessage: 'Error while checking if content is ingested',
        status: 'failed'
      };
    }
  }

  private async checkUrlToItemResolverConfigured(id: string): Promise<CheckResult> {
    const externalConnection: ExternalConnectors.ExternalConnection = this.checksStatus.find(r => r.id === 'loadExternalConnection')!.data;

    if (!externalConnection.activitySettings?.urlToItemResolvers?.some(r => r)) {
      return {
        id,
        errorMessage: 'urlToItemResolver is not configured',
        status: 'failed'
      };
    }

    return {
      id,
      status: 'passed'
    };
  }

  private async checkSemanticLabels(id: string): Promise<CheckResult> {
    const schema: ExternalConnectors.Schema = this.checksStatus.find(r => r.id === 'loadSchema')!.data;
    const hasLabels = schema.properties?.some(p => p.labels?.some(l => l));

    if (!hasLabels) {
      return {
        id,
        errorMessage: `Schema does not have semantic labels`,
        status: 'failed'
      };
    }

    return {
      id,
      status: 'passed'
    };
  }

  private async checkResultType(id: string): Promise<CheckResult> {
    const externalConnection: ExternalConnectors.ExternalConnection = this.checksStatus.find(r => r.id === 'loadExternalConnection')!.data;
    if (!externalConnection.searchSettings?.searchResultTemplates?.some(t => t)) {
      return {
        id,
        errorMessage: `Connection has no result types`,
        status: 'failed'
      };
    }

    return {
      id,
      status: 'passed'
    };
  }
}

export default new ExternalConnectionDoctorCommand();