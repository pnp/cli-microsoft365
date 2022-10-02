import * as chalk from 'chalk';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSidePageProperties } from './ClientSidePageProperties';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  sourceName: string;
  targetUrl: string;
  webUrl: string;
  overwrite?: boolean;
}

class SpoPageCopyCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_COPY}`;
  }

  public get description(): string {
    return 'Creates a copy of a modern page or template';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'PageLayoutType', 'Title', 'Url'];
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
        overwrite: !!args.options.overwrite
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--sourceName <sourceName>'
      },
      {
        option: '--targetUrl <targetUrl>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--overwrite'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let { webUrl } = args.options;
    const { targetUrl, overwrite } = args.options;
    webUrl = this.removeTrailingSlash(webUrl);

    let sourceFullName: string = args.options.sourceName.toLowerCase();
    const targetPageInfo: { siteUrl: string, pageName: string } = this.getTargetSiteUrl(webUrl, targetUrl.toLowerCase());
    let { siteUrl: targetSiteUrl, pageName: targetFullName } = targetPageInfo;

    if (sourceFullName.indexOf('.aspx') < 0) {
      sourceFullName += '.aspx';
    }
    if (targetFullName.indexOf('.aspx') < 0) {
      targetFullName += '.aspx';
    }

    targetSiteUrl = this.removeTrailingSlash(targetSiteUrl);
    if (targetFullName.startsWith('/')) {
      targetFullName = targetFullName.substring(1);
    }

    if (this.verbose) {
      logger.logToStderr(`Creating page copy...`);
    }

    try {
      let requestOptions: any = {
        url: `${webUrl}/_api/SP.MoveCopyUtil.CopyFileByPath()`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          srcPath: { DecodedUrl: `${webUrl}/sitepages/${sourceFullName}` },
          destPath: { DecodedUrl: `${targetSiteUrl}/sitepages/${targetFullName}` },
          options: { ResetAuthorAndCreatedOnCopy: true, ShouldBypassSharedLocks: true },
          overwrite: !!overwrite
        },
        responseType: 'json'
      };

      await request.post(requestOptions);

      requestOptions = {
        url: `${targetSiteUrl}/_api/sitepages/pages/GetByUrl('sitepages/${targetFullName}')`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<ClientSidePageProperties>(requestOptions);
      logger.log(res);

      if (this.verbose) {
        logger.logToStderr(chalk.green('DONE'));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getTargetSiteUrl(webUrl: string, targetFullName: string): { siteUrl: string, pageName: string } {
    const siteSplit: string[] = targetFullName.split('sitepages/');

    if (targetFullName.startsWith("http")) {
      return {
        siteUrl: siteSplit[0],
        pageName: siteSplit[1]
      };
    }
    else {
      return {
        siteUrl: webUrl,
        pageName: siteSplit.length > 1 ? siteSplit[1] : targetFullName
      };
    }
  }

  private removeTrailingSlash(value: string) {
    if (value.endsWith('/')) {
      value = value.substring(0, value.length - 1);
    }

    return value;
  }
}

module.exports = new SpoPageCopyCommand(); 