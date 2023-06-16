import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
  confirm?: boolean;
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
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.removePage(logger, args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>(
        {
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove the page '${args.options.name}'?`
        });

      if (result.continue) {
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
        logger.logToStderr(`Removing page ${pageName}...`);
      }

      const requestOptions: any = {
        url: `${args.options
          .webUrl}/_api/web/getfilebyserverrelativeurl('${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages/${pageName}')`,
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

module.exports = new SpoPageRemoveCommand();
