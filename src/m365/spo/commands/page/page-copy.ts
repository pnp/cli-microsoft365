import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.overwrite = !!args.options.overwrite;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let { webUrl, targetUrl, overwrite } = args.options;
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

    const requestOptions: any = {
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

    request
      .post(requestOptions)
      .then((): Promise<ClientSidePageProperties> => {
        const requestOptions: any = {
          url: `${targetSiteUrl}/_api/sitepages/pages/GetByUrl('sitepages/${targetFullName}')`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get<ClientSidePageProperties>(requestOptions)
      })
      .then((res: ClientSidePageProperties): void => {
        logger.log(res);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      })
      .catch(err => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--sourceName <sourceName>',
        description: 'The name of the source file'
      },
      {
        option: '--targetUrl <targetUrl>',
        description: 'The url of the target file. You are able to provide the page its name, relative path, or absolute path'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to copy is located'
      },
      {
        option: '--overwrite',
        description: 'Overwrite the target page when it already exists'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
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