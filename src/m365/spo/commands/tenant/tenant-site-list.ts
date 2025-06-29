import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { TenantSiteProperties } from './TenantSiteProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
  webTemplate?: string;
  filter?: string;
  includeOneDriveSites?: boolean;
}

class SpoTenantSiteListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_SITE_LIST;
  }

  public get description(): string {
    return 'Lists sites of the given type';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url'];
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
        webTemplate: args.options.webTemplate,
        type: args.options.type,
        filter: (!(!args.options.filter)).toString(),
        includeOneDriveSites: typeof args.options.includeOneDriveSites !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type [type]',
        autocomplete: ['TeamSite', 'CommunicationSite']
      },
      {
        option: '--webTemplate [webTemplate]'
      },
      {
        option: '--filter [filter]'
      },
      {
        option: '--includeOneDriveSites'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type && args.options.webTemplate) {
          return 'Specify either type or webTemplate, but not both';
        }

        const typeValues = ['TeamSite', 'CommunicationSite'];
        if (args.options.type &&
          typeValues.indexOf(args.options.type) < 0) {
          return `${args.options.type} is not a valid value for the type option. Allowed values are ${typeValues.join('|')}`;
        }

        if (args.options.includeOneDriveSites
          && (args.options.type || args.options.webTemplate)) {
          return 'When using includeOneDriveSites, don\'t specify the type or webTemplate options';
        }

        return true;
      }
    );
  }

  public alias(): string[] | undefined {
    return [commands.SITE_LIST];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const webTemplate: string = this.getWebTemplateId(args.options);
    const includeOneDriveSites: boolean = args.options.includeOneDriveSites || false;

    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving list of site collections...`);
      }

      const allSites: TenantSiteProperties[] = await spo.getAllSites(spoAdminUrl, logger, this.verbose, formatting.escapeXml(args.options.filter || ''), includeOneDriveSites, webTemplate);
      await logger.log(allSites);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getWebTemplateId(options: Options): string {
    if (options.webTemplate) {
      return options.webTemplate;
    }

    if (options.includeOneDriveSites) {
      return '';
    }

    switch (options.type) {
      case "TeamSite":
        return 'GROUP#0';
      case "CommunicationSite":
        return 'SITEPAGEPUBLISHING#0';
      default:
        return '';
    }
  }
}

export default new SpoTenantSiteListCommand();