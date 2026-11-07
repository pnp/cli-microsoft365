import chalk from 'chalk';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ClientSidePageProperties } from './ClientSidePageProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  sourceName: z.string(),
  targetUrl: z.string(),
  overwrite: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageCopyCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_COPY}`;
  }

  public get description(): string {
    return 'Creates a copy of a modern page or template';
  }

  public get schema(): z.ZodType {
    return options;
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
      await logger.logToStderr(`Creating page copy...`);
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
      await logger.log(res);

      if (this.verbose) {
        await logger.logToStderr(chalk.green('DONE'));
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

  private removeTrailingSlash(value: string): string {
    if (value.endsWith('/')) {
      value = value.substring(0, value.length - 1);
    }

    return value;
  }
}

export default new SpoPageCopyCommand(); 