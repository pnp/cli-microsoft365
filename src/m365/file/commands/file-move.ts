import { Drive } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../cli/Logger.js';
import { globalOptionsZod } from '../../../Command.js';
import GraphCommand from '../../base/GraphCommand.js';
import { setTimeout } from 'timers/promises';
import commands from '../commands.js';
import request, { CliRequestOptions } from '../../../request.js';
import { spo } from '../../../utils/spo.js';
import { urlUtil } from '../../../utils/urlUtil.js';
import { drive } from '../../../utils/drive.js';
import { validation } from '../../../utils/validation.js';

const nameConflictBehaviorOptions = ['fail', 'replace', 'rename'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  sourceUrl: z.string().alias('s'),
  targetUrl: z.string().alias('t'),
  newName: z.string().optional(),
  nameConflictBehavior: z.enum(nameConflictBehaviorOptions).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class FileMoveCommand extends GraphCommand {
  private pollingInterval: number = 10_000;

  public get name(): string {
    return commands.MOVE;
  }

  public get description(): string {
    return 'Moves a file to another location using the Microsoft Graph';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const { webUrl, sourceUrl, targetUrl, nameConflictBehavior, newName, verbose } = args.options;
      const sourcePath: string = this.getAbsoluteUrl(webUrl, sourceUrl);
      const destinationPath: string = this.getAbsoluteUrl(webUrl, targetUrl);

      const { driveId, itemId } = await this.getDriveIdAndItemId(webUrl, sourcePath, sourceUrl, logger, verbose);
      const targetSiteUrl: string = urlUtil.getTargetSiteAbsoluteUrl(webUrl, targetUrl);
      const targetFolderUrl: string = this.getAbsoluteUrl(targetSiteUrl, targetUrl);
      const { driveId: targetDriveId, itemId: targetItemId } = await this.getDriveIdAndItemId(targetSiteUrl, targetFolderUrl, targetUrl, logger, verbose);

      const requestOptions: CliRequestOptions = this.getRequestOptions(driveId, itemId, targetDriveId, targetItemId, newName, sourcePath, nameConflictBehavior);

      if (verbose) {
        await logger.logToStderr(`Moving file '${sourcePath}' to '${destinationPath}'...`);
      }

      if (driveId === targetDriveId) {
        await request.patch(requestOptions);
      }
      else {
        const response: any = await request.post(requestOptions);
        await this.waitUntilCopyOperationCompleted(response.headers.location, logger);

        const itemUrl = `${this.resource}/v1.0/drives/${driveId}/items/${itemId}`;
        await request.delete({ url: itemUrl, headers: requestOptions.headers });
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getAbsoluteUrl(webUrl: string, url: string): string {
    return url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
  }

  private async getDriveIdAndItemId(webUrl: string, folderUrl: string, sourceUrl: string, logger: Logger, verbose?: boolean): Promise<{ driveId: string, itemId: string }> {
    const siteId: string = await spo.getSiteIdByMSGraph(webUrl, logger, verbose);
    const driveDetails: Drive = await drive.getDriveByUrl(siteId, new URL(folderUrl), logger, verbose);
    const itemId: string = await drive.getDriveItemId(driveDetails, new URL(folderUrl), logger, verbose);
    return { driveId: driveDetails.id!, itemId };
  }

  private getRequestOptions(sourceDriveId: string, sourceItemId: string, targetDriveId: string, targetItemId: string, newName: string | undefined, sourcePath: string, nameConflictBehavior: string | undefined): CliRequestOptions {
    const apiUrl =
      sourceDriveId === targetDriveId
        ? `${this.resource}/v1.0/drives/${sourceDriveId}/items/${sourceItemId}`
        : `${this.resource}/v1.0/drives/${sourceDriveId}/items/${sourceItemId}/copy`;

    const queryParameters: string = nameConflictBehavior && nameConflictBehavior !== 'fail'
      ? `@microsoft.graph.conflictBehavior=${nameConflictBehavior}`
      : '';
    const urlWithQuery = `${apiUrl}${queryParameters ? `?${queryParameters}` : ''}`;

    const requestOptions: CliRequestOptions = {
      url: urlWithQuery,
      headers: { accept: 'application/json;odata.metadata=none' },
      responseType: 'json',
      fullResponse: true,
      data: { parentReference: { driveId: targetDriveId, id: targetItemId } }
    };

    if (newName) {
      const sourceFileName = sourcePath.substring(sourcePath.lastIndexOf('/') + 1);
      const sourceFileExtension = sourceFileName.includes('.') ? sourceFileName.substring(sourceFileName.lastIndexOf('.')) : '';
      const newNameExtension = newName.includes('.') ? newName.substring(newName.lastIndexOf('.')) : '';
      requestOptions.data.name = newNameExtension ? `${newName.replace(newNameExtension, '')}${sourceFileExtension}` : `${newName}${sourceFileExtension}`;
    }

    return requestOptions;
  }

  private async waitUntilCopyOperationCompleted(monitorUrl: string, logger: Logger): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: monitorUrl,
      responseType: 'json'
    };

    const response: any = await request.get(requestOptions);
    if (response.status === 'completed') {
      if (this.verbose) {
        await logger.logToStderr('Copy operation completed succesfully. Returning...');
      }
      return;
    }
    else if (response.status === 'failed') {
      throw response.error.message;
    }
    else {
      if (this.verbose) {
        await logger.logToStderr(`Still copying. Retrying in ${this.pollingInterval / 1000} seconds...`);
      }
      await setTimeout(this.pollingInterval);
      await this.waitUntilCopyOperationCompleted(monitorUrl, logger);
    }
  }
}

export default new FileMoveCommand();
