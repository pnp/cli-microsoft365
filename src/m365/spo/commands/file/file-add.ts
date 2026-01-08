import fs from 'fs';
import path from 'path';
import { v4 } from 'uuid';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { fsUtil } from '../../../../utils/fsUtil.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

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
  overwrite?: boolean;
  fileName?: string;
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
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        contentType: typeof args.options.contentType !== 'undefined',
        checkOut: !!args.options.checkOut,
        checkInComment: typeof args.options.checkInComment !== 'undefined',
        approve: !!args.options.approve,
        approveComment: typeof args.options.approveComment !== 'undefined',
        publish: !!args.options.publish,
        publishComment: typeof args.options.publishComment !== 'undefined',
        overwrite: !!args.options.overwrite,
        fileName: typeof args.options.fileName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--folder <folder>'
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
      },
      {
        option: '--overwrite [overwrite]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--fileName [fileName]'
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

  #initTypes(): void {
    this.types.string.push('webUrl', 'folder', 'path', 'contentType', 'checkInComment', 'approveComment', 'publishComment', 'fileName');
    this.types.boolean.push('checkOut', 'approve', 'publish', 'overwrite');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const folderPath: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folder);
    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = fsUtil.getSafeFileName(args.options.fileName ?? path.basename(fullPath));

    let isCheckedOut: boolean = false;
    let listSettings: ListSettings;

    if (this.verbose) {
      await logger.logToStderr(`file name: ${fileName}...`);
      await logger.logToStderr(`folder path: ${folderPath}...`);
    }

    if (args.options.overwrite === undefined) {
      await this.warn(logger, 'In the next major version, the --overwrite option will default to false. To avoid this warning, please set the --overwrite option explicitly to true or false.');
    }

    try {
      if (args.options.overwrite === false) {
        try {
          const requestOptions: CliRequestOptions = {
            url: `${args.options.webUrl}/_api/Web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderPath + '/' + fileName)}')?$select=Exists`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };
          await request.get<void>(requestOptions);

          throw `File '${fileName}' already exists in folder '${folderPath}'. To overwrite the file, use the --overwrite option.`;
        }
        catch (err: any) {
          if (typeof err === 'string') {
            throw err;
          }

          if (this.verbose) {
            await logger.logToStderr(`File '${fileName}' does not exist in folder '${folderPath}'. Proceeding with upload.`);
          }
        }
      }

      try {
        if (this.verbose) {
          await logger.logToStderr('Check if the specified folder exists.');
        }

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderPath)}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
        await request.get<void>(requestOptions);
      }
      catch {
        // folder does not exist so will attempt to create the folder tree
        await spo.ensureFolder(args.options.webUrl, folderPath, logger, this.verbose);
      }

      if (args.options.checkOut) {
        await this.fileCheckOut(fileName, args.options.webUrl, folderPath);
        // flag the file is checkedOut by the command
        // so in case of command failure we can try check it in
        isCheckedOut = true;
      }

      if (this.verbose) {
        await logger.logToStderr(`Upload file to site ${args.options.webUrl}...`);
      }

      const fileStats: fs.Stats = fs.statSync(fullPath);
      const fileSize: number = fileStats.size;
      if (this.verbose) {
        await logger.logToStderr(`File size is ${fileSize} bytes`);
      }

      let fileUploadResult: any;
      // only up to 250 MB are allowed in a single request
      if (fileSize > this.fileChunkingThreshold) {
        const fileChunkCount: number = Math.ceil(fileSize / this.fileChunkSize);
        if (this.verbose) {
          await logger.logToStderr(`Uploading ${fileSize} bytes in ${fileChunkCount} chunks...`);
        }

        // initiate chunked upload session
        const uploadId: string = v4();
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderPath)}')/Files/GetByPathOrAddStub(DecodedUrl='${formatting.encodeQueryParameter(fileName)}')/StartUpload(uploadId=guid'${uploadId}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
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
          fileUploadResult = await this.uploadFileChunks(fileUploadInfo, logger);

          if (this.verbose) {
            await logger.logToStderr(`Finished uploading ${fileUploadInfo.Position} bytes in ${fileChunkCount} chunks`);
          }
        }
        catch (err: any) {
          if (this.verbose) {
            await logger.logToStderr('Cancelling upload session due to error...');
          }

          const requestOptions: CliRequestOptions = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/cancelupload(uploadId=guid'${uploadId}')`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          try {
            await request.post<void>(requestOptions);
            throw err;
          }
          catch (err: any) {
            if (this.verbose) {
              await logger.logToStderr(`Failed to cancel upload session: ${err}`);
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
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderPath)}')/Files/Add(url='${formatting.encodeQueryParameter(fileName)}', overwrite=true)`,
          data: fileBody,
          headers: {
            accept: 'application/json;odata=nometadata',
            'content-length': bodyLength
          },
          maxBodyLength: this.fileChunkingThreshold,
          responseType: 'json'
        };

        fileUploadResult = await request.post<any>(requestOptions);
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
        // perform list item update and check-in
        await this.validateUpdateListItem(args.options.webUrl, folderPath, fileName, fieldsToUpdate, logger, args.options.checkInComment);
      }
      else if (isCheckedOut) {
        // perform check-in
        await this.fileCheckIn(args.options.webUrl, folderPath, fileName, args.options.checkInComment);
      }

      // approve and publish cannot be used together
      // when approve is used it will automatically publish the file
      // so then no need to publish afterwards
      if (args.options.approve) {
        if (this.verbose) {
          await logger.logToStderr(`Approve file ${fileName}`);
        }

        // approve the existing file with given comment
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/approve(comment='${formatting.encodeQueryParameter(args.options.approveComment || '')}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
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
          await logger.logToStderr(`Publish file ${fileName}`);
        }

        // publish the existing file with given comment
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/publish(comment='${formatting.encodeQueryParameter(args.options.publishComment || '')}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }

      await logger.log(fileUploadResult);
    }
    catch (err: any) {
      if (isCheckedOut) {
        // in a case the command has done checkout
        // then have to rollback the checkout

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/UndoCheckOut()`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        try {
          await request.post(requestOptions);
        }
        catch (err: any) {
          if (this.verbose) {
            await logger.logToStderr('Could not rollback file checkout');
            await logger.logToStderr(err);
          }
        }
      }

      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async listHasContentType(contentType: string, webUrl: string, listSettings: ListSettings, logger: any): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Getting list of available content types ...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/lists('${listSettings.Id}')/contenttypes?$select=Name,Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
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
      url: `${webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folder)}')/Files('${formatting.encodeQueryParameter(fileName)}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.get<void>(requestOptionsGetFile);

    const requestOptionsCheckOut: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folder)}')/Files('${formatting.encodeQueryParameter(fileName)}')/CheckOut()`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post<void>(requestOptionsCheckOut);
  }

  private async uploadFileChunks(info: FileUploadInfo, logger: Logger): Promise<any> {
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
        fileBuffer = fileBuffer.subarray(0, readCount);
      }

      const requestOptions: CliRequestOptions = {
        url: `${info.WebUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(info.FolderPath)}')/Files('${formatting.encodeQueryParameter(info.Name)}')/${isLastChunk ? 'Finish' : 'Continue'}Upload(uploadId=guid'${info.Id}',fileOffset=${offset})`,
        data: fileBuffer,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-length': readCount
        },
        maxBodyLength: this.fileChunkingThreshold,
        responseType: 'json'
      };

      try {
        const uploadResponse = await request.post<void>(requestOptions);
        if (this.verbose) {
          await logger.logToStderr(`Uploaded ${info.Position} of ${info.Size} bytes (${Math.round(100 * info.Position / info.Size)}%)`);
        }

        if (isLastChunk) {
          return uploadResponse;
        }
        else {
          return this.uploadFileChunks(info, logger);
        }
      }
      catch (err: any) {
        if (--info.RetriesLeft > 0) {
          if (this.verbose) {
            await logger.logToStderr(`Retrying to upload chunk due to error: ${err}`);
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
        catch {
          // Do nothing
        }
      }

      if (--info.RetriesLeft > 0) {
        if (this.verbose) {
          await logger.logToStderr(`Retrying to read chunk due to error: ${err}`);
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
      await logger.logToStderr(`Getting list details in order to get its available content types afterwards...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folder)}')/Files('${formatting.encodeQueryParameter(fileName)}')/ListItemAllFields/ParentList?$Select=Id,EnableModeration,EnableVersioning,EnableMinorVersions`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  private async validateUpdateListItem(webUrl: string, folderPath: string, fileName: string, fieldsToUpdate: FieldValue[], logger: any, checkInComment?: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Validate and update list item values for file ${fileName}`);
    }

    const requestBody: any = {
      formValues: fieldsToUpdate,
      bNewDocumentUpdate: true, // true = will automatically check-in the item, but we will use it to perform system update and also do a check-in
      checkInComment: checkInComment || ''
    };

    if (this.verbose) {
      await logger.logToStderr('ValidateUpdateListItem will perform the check-in ...');
    }

    // update the existing file list item fields
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderPath)}')/Files('${formatting.encodeQueryParameter(fileName)}')/ListItemAllFields/ValidateUpdateListItem()`,
      headers: {
        accept: 'application/json;odata=nometadata'
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

  private async fileCheckIn(webUrl: string, folderUrl: string, fileName: string, checkInComment?: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(folderUrl)}')/Files('${formatting.encodeQueryParameter(fileName)}')/CheckIn(comment='${formatting.encodeQueryParameter(checkInComment || '')}',checkintype=0)`,
      headers: {
        accept: 'application/json;odata=nometadata'
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
      'overwrite',
      'fileName',
      'debug',
      'verbose',
      'output',
      'query',
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

export default new SpoFileAddCommand();
