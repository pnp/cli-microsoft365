import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SiteDesign } from './SiteDesign.js';
import { getBuiltInSiteDesignTemplateName } from './BuiltInSiteDesigns.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  builtIn?: boolean;
}

class SpoSiteDesignListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_LIST;
  }

  public get description(): string {
    return 'Lists available site designs for creating modern sites';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'IsDefault', 'Title', 'Version', 'WebTemplate', 'Template'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        builtIn: args.options.builtIn || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-b, --builtIn'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.builtIn) {
        requestOptions.headers['content-type'] = 'application/json;charset=utf-8';
        requestOptions.data = { store: 1 };
      }

      const res: { value: SiteDesign[] } = await request.post(requestOptions);

      if (args.options.builtIn) {
        await logger.log(res.value.map(siteDesign => ({
          ...siteDesign,
          Template: getBuiltInSiteDesignTemplateName(siteDesign.Id)
        })));
      }
      else {
        await logger.log(res.value);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteDesignListCommand();
