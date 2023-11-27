import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
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

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-u, --webUrl <webUrl>'
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removePage(logger, args);
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to remove the page '${args.options.name}'?` });

      if (result) {
        await this.removePage(logger, args);
      }
    }
  }

  private async removePage(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let requestDigest: string = '';
      let pageName: string = args.options.name;

      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      requestDigest = reqDigest.FormDigestValue;

      if (!pageName.endsWith('.aspx')) {
        pageName += '.aspx';
      }

      if (this.verbose) {
        await logger.logToStderr(`Removing page ${pageName}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options
          .webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages/${pageName}')`,
        headers: {
          'X-RequestDigest': requestDigest,
          'X-HTTP-Method': 'DELETE',
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageRemoveCommand();
