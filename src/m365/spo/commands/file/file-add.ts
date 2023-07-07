import * as fs from 'fs';
import * as path from 'path';
import { v4 } from 'uuid';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { fsUtil } from '../../../../utils/fsUtil';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folder: string;
  path: string;
  contentType?: string;
  checkOut?: boolean;
  checkInComment?: string;
  approve?: boolean;
  approveComment?: string;
  publish?: boolean;
  publishComment?: string;
}

interface FieldValue {
  FieldName: string;
  FieldValue: any;
}

interface FieldValueResult extends FieldValue {
  ErrorMessage: string;
  HasException: boolean;
  ItemId: number;
}

interface ListSettings {
  Id: string;
  EnableVersioning: boolean;
  EnableModeration: boolean;
  EnableMinorVersions: boolean;
}

interface FileUploadInfo {
  Name: string;
  FilePath: string;
  WebUrl: string;
  FolderPath: string;
  Id: string;
  RetriesLeft: number;
  Size: number;
  Position: number;
}

class SpoFileAddCommand extends SpoCommand {
  private readonly fileChunkingThreshold: number = 250 * 1024 * 1024;  // max 250 MB
  private readonly fileChunkSize: number = 250 * 1024 * 1024;  // max fileChunkingThreshold
  private readonly fileChunkRetryAttempts: number = 5;

  public get name(): string {
    return commands.FILE_ADD;
  }

  public get description(): string {
    return 'Uploads file to the specified folder';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
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
        contentType: (!(!args.options.contentType)).toString(),
        checkOut: args.options.checkOut || false,
        checkInComment: (!(!args.options.checkInComment)).toString(),
        approve: args.options.approve || false,
        approveComment: (!(!args.options.approveComment)).toString(),
        publish: args.options.publish || false,
        publishComment: (!(!args.options.publishComment)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-f, --folder <folder>'
      },
      {
        option: '-p, --path <path>'
      },
      {
        option: '-c, --contentType [contentType]'
      },
      {
        option: '--checkOut'
      },
      {
        option: '--checkInComment [checkInComment]'
      },
      {
        option: '--approve'
      },
      {
        option: '--approveComment [approveComment]'
      },
      {
        option: '--publish'
      },
      {
        option: '--publishComment [publishComment]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.path && !fs.existsSync(args.options.path)) {
          return 'Specified path of the file to add does not exist';
        }

        if (args.options.publishComment && !args.options.publish) {
          return '--publishComment cannot be used without --publish';
        }

        if (args.options.approveComment && !args.options.approve) {
          return '--approveComment cannot be used without --approve';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const folderPath: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folder);
    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = fsUtil.getSafeFileName(path.basename(fullPath));

    let isCheckedOut: boolean = false;
    let listSettings: ListSettings;

    if (this.debug) {
      logger.logToStderr(`folder path: ${folderPath}...`);
    }

    if (this.debug) {
      logger.logToStderr('Check if the specified folder exists.');
      logger.logToStderr('');
    }

    if (this.debug) {
      logger.logToStderr(`file name: ${fileName}...`);
    }

    try {
      try {
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderPath)}')`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          }
        };
        await request.get<void>(requestOptions);
      }
      catch (err: any) {
        // folder does not exist so will attempt to create the folder tree
        await spo.ensureFolder(args.options.webUrl, folderPath, logger, this.debug);
      }

      if (args.options.checkOut) {
        await this.fileCheckOut(fileName, args.options.webUrl, folderPath);
        // flag the file is checkedOut by the command
        // so in case of command failure we can try check it in
        isCheckedOut = true;
      }

      if (this.verbose) {
        logger.logToStderr(`Upload file to site ${args.options.webUrl}...`);
      }

      const fileStats: fs.Stats = fs.statSync(fullPath);
      const fileSize: number = fileStats.size;
      if (this.debug) {
        logger.logToStderr(`File size is ${fileSize} bytes`);
      }

      // only up to 250 MB are allowed in a single request
      if (fileSize > this.fileChunkingThreshold) {
        const fileChunkCount: number = Math.ceil(fileSize / this.fileChunkSize);
        if (this.verbose) {
          logger.logToStderr(`Uploading ${fileSize} bytes in ${fileChunkCount} chunks...`);
        }

        // initiate chunked upload session
        const uploadId: string = v4();
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderPath)}')/Files/GetByPathOrAddStub(DecodedUrl='${formatting.encodeQueryParameter(fileName)}')/StartUpload(uploadId=guid'${uploadId}')`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          }
        };

        await request.post<void>(requestOptions);
        // session started successfully, now upload our file chunks
        const fileUploadInfo: FileUploadInfo = {
          Name: fileName,
          FilePath: fullPath,
          WebUrl: args.options.webUrl,
          FolderPath: folderPath,
          Id: uploadId,
          RetriesLeft: this.fileChunkRetryAttempts,
          Position: 0,
          Size: fileSize
        };

        try {
          await this.uploadFileChunks(fileUploadInfo, logger);

          if (this.verbose) {
            logger.logToStderr(`Finished uploading ${fileUploadInfo.Position} bytes in ${fileChunkCount} chunks`);
          }
        }
        catch (err: any) {
          if (this.verbose) {
            logger.logToStderr('Cancelling upload session due to error...');
          }

          const requestOptions: CliRequestOptions = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/cancelupload(uploadId=guid'${uploadId}')`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            }
          };

          try {
            await request.post<void>(requestOptions);
            throw err;
          }
          catch (err: any) {
            if (this.debug) {
              logger.logToStderr(`Failed to cancel upload session: ${err}`);
            }
            throw err;
          }
        }
      }
      else {
        // upload small file in a single request
        const fileBody: Buffer = fs.readFileSync(fullPath);
        const bodyLength: number = fileBody.byteLength;

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderPath)}')/Files/Add(url='${formatting.encodeQueryParameter(fileName)}', overwrite=true)`,
          data: fileBody,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'content-length': bodyLength
          },
          maxBodyLength: this.fileChunkingThreshold
        };

        await request.post(requestOptions);
      }

      if (args.options.contentType || args.options.publish || args.options.approve) {
        listSettings = await this.getFileParentList(fileName, args.options.webUrl, folderPath, logger);

        if (args.options.contentType) {
          await this.listHasContentType(args.options.contentType, args.options.webUrl, listSettings, logger);
        }
      }

      // check if there are unknown options
      // and map them as fields to update
      const fieldsToUpdate: FieldValue[] = this.mapUnknownOptionsAsFieldValue(args.options);

      if (args.options.contentType) {
        fieldsToUpdate.push({
          FieldName: 'ContentType',
          FieldValue: args.options.contentType
        });
      }

      if (fieldsToUpdate.length > 0) {
        // perform list item update and checkin
        await this.validateUpdateListItem(args.options.webUrl, folderPath, fileName, fieldsToUpdate, logger, args.options.checkInComment);
      }
      else if (isCheckedOut) {
        // perform checkin
        await this.fileCheckIn(args, fileName);
      }

      // approve and publish cannot be used together
      // when approve is used it will automatically publish the file
      // so then no need to publish afterwards
      if (args.options.approve) {
        if (this.verbose) {
          logger.logToStderr(`Approve file ${fileName}`);
        }

        // approve the existing file with given comment
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/approve(comment='${formatting.encodeQueryParameter(args.options.approveComment || '')}')`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      else if (args.options.publish) {
        if (listSettings!.EnableModeration && listSettings!.EnableMinorVersions) {
          throw 'The file cannot be published without approval. Moderation for this list is enabled. Use the --approve option instead of --publish to approve and publish the file';
        }

        if (this.verbose) {
          logger.logToStderr(`Publish file ${fileName}`);
        }

        // publish the existing file with given comment
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/publish(comment='${formatting.encodeQueryParameter(args.options.publishComment || '')}')`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
    }
    catch (err: any) {
      if (isCheckedOut) {
        // in a case the command has done checkout
        // then have to rollback the checkout

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/UndoCheckOut()`
        };

        try {
          await request.post(requestOptions);
        }
        catch (err: any) {
          if (this.verbose) {
            logger.logToStderr('Could not rollback file checkout');
            logger.logToStderr(err);
            logger.logToStderr('');
          }
        }
      }

      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async listHasContentType(contentType: string, webUrl: string, listSettings: ListSettings, logger: any): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Getting list of available content types ...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/lists('${listSettings.Id}')/contenttypes?$select=Name,Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    const response = await request.get<any>(requestOptions);
    // check if the specified content type is in the list
    for (const ct of response.value) {
      if (ct.Id.StringValue === contentType || ct.Name === contentType) {
        return;
      }
    }

    throw `Specified content type '${contentType}' doesn't exist on the target list`;
  }

  private async fileCheckOut(fileName: string, webUrl: string, folder: string): Promise<void> {
    // check if file already exists, otherwise it can't be checked out
    const requestOptionsGetFile: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files('${formatting.encodeQueryParameter(fileName)}')`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      }
    };

    await request.get<void>(requestOptionsGetFile);

    const requestOptionsCheckOut: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files('${formatting.encodeQueryParameter(fileName)}')/CheckOut()`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post<void>(requestOptionsCheckOut);
  }

  private async uploadFileChunks(info: FileUploadInfo, logger: any): Promise<void> {
    let fd: number = 0;
    try {
      fd = fs.openSync(info.FilePath, 'r');
      let fileBuffer: Buffer = Buffer.alloc(this.fileChunkSize);
      const readCount: number = fs.readSync(fd, fileBuffer, 0, this.fileChunkSize, info.Position);
      fs.closeSync(fd);
      fd = 0;

      const offset: number = info.Position;
      info.Position += readCount;
      const isLastChunk: boolean = info.Position >= info.Size;
      if (isLastChunk) {
        // trim buffer for last chunk
        fileBuffer = fileBuffer.slice(0, readCount);
      }

      const requestOptions: CliRequestOptions = {
        url: `${info.WebUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(info.FolderPath)}')/Files('${formatting.encodeQueryParameter(info.Name)}')/${isLastChunk ? 'Finish' : 'Continue'}Upload(uploadId=guid'${info.Id}',fileOffset=${offset})`,
        data: fileBuffer,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-length': readCount
        },
        maxBodyLength: this.fileChunkingThreshold
      };

      try {
        await request.post<void>(requestOptions);
        if (this.verbose) {
          logger.logToStderr(`Uploaded ${info.Position} of ${info.Size} bytes (${Math.round(100 * info.Position / info.Size)}%)`);
        }

        if (isLastChunk) {
          return;
        }
        else {
          return this.uploadFileChunks(info, logger);
        }
      }
      catch (err: any) {
        if (--info.RetriesLeft > 0) {
          if (this.verbose) {
            logger.logToStderr(`Retrying to upload chunk due to error: ${err}`);
          }
          info.Position -= readCount;  // rewind
          return this.uploadFileChunks(info, logger);
        }
        else {
          throw err;
        }
      }
    }
    catch (err) {
      if (fd) {
        try {
          fs.closeSync(fd);
          /* c8 ignore next */
        }
        catch { }
      }

      if (--info.RetriesLeft > 0) {
        if (this.verbose) {
          logger.logToStderr(`Retrying to read chunk due to error: ${err}`);
        }
        return this.uploadFileChunks(info, logger);
      }
      else {
        throw err;
      }
    }
  }

  private async getFileParentList(fileName: string, webUrl: string, folder: string, logger: any): Promise<ListSettings> {
    if (this.verbose) {
      logger.logToStderr(`Getting list details in order to get its available content types afterwards...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folder)}')/Files('${formatting.encodeQueryParameter(fileName)}')/ListItemAllFields/ParentList?$Select=Id,EnableModeration,EnableVersioning,EnableMinorVersions`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  private async validateUpdateListItem(webUrl: string, folderPath: string, fileName: string, fieldsToUpdate: FieldValue[], logger: any, checkInComment?: string): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Validate and update list item values for file ${fileName}`);
    }

    const requestBody: any = {
      formValues: fieldsToUpdate,
      bNewDocumentUpdate: true, // true = will automatically checkin the item, but we will use it to perform system update and also do a checkin
      checkInComment: checkInComment || ''
    };

    if (this.debug) {
      logger.logToStderr('ValidateUpdateListItem will perform the checkin ...');
      logger.logToStderr('');
    }

    // update the existing file list item fields
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/ListItemAllFields/ValidateUpdateListItem()`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    const res = await request.post<any>(requestOptions);
    // check for field value update for errors
    const fieldValues: FieldValueResult[] = res.value;
    for (const fieldValue of fieldValues) {
      if (fieldValue.HasException) {
        throw `Update field value error: ${JSON.stringify(fieldValues)}`;
      }
    }
    return;
  }

  private async fileCheckIn(args: any, fileName: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${formatting.encodeQueryParameter(args.options.folder)}')/Files('${formatting.encodeQueryParameter(fileName)}')/CheckIn(comment='${formatting.encodeQueryParameter(args.options.checkInComment || '')}',checkintype=0)`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private mapUnknownOptionsAsFieldValue(options: Options): FieldValue[] {
    const result: any = [];
    const excludeOptions: string[] = [
      'webUrl',
      'folder',
      'path',
      'contentType',
      'checkOut',
      'checkInComment',
      'approve',
      'approveComment',
      'publish',
      'publishComment',
      'debug',
      'verbose',
      'output',
      '_',
      'u',
      'p',
      'f',
      'o',
      'c'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        result.push({ FieldName: key, FieldValue: (<any>options)[key].toString() });
      }
    });

    return result;
  }
}

module.exports = new SpoFileAddCommand();
