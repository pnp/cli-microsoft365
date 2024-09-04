import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  name: string;
  recycle?: boolean;
  bypassSharedLock?: boolean;
  force?: boolean;
}

class SpoPageRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a modern page';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force,
        recycle: !!args.options.recycle,
        bypassSharedLock: !!args.options.bypassSharedLock
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '--recycle'
      },
      {
        option: '--bypassSharedLock'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  #initTypes(): void {
    this.types.string.push('name', 'webUrl');
    this.types.boolean.push('force', 'bypassSharedLock', 'recycle');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removePage(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove page '${args.options.name}'?` });

      if (result) {
        await this.removePage(logger, args);
      }
    }
  }

  private async removePage(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      // Remove leading slashes from the page name (page can be nested in folders)
      let pageName: string = urlUtil.removeLeadingSlashes(args.options.name);
      if (!pageName.toLowerCase().endsWith('.aspx')) {
        pageName += '.aspx';
      }

      if (this.verbose) {
        await logger.logToStderr(`Removing page ${pageName}...`);
      }

      const filePath = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${pageName}`;
      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(filePath)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.bypassSharedLock) {
        requestOptions.headers!.Prefer = 'bypass-shared-lock';
      }
      if (args.options.recycle) {
        requestOptions.url += '/Recycle';

        await request.post(requestOptions);
      }
      else {
        await request.delete(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageRemoveCommand();
