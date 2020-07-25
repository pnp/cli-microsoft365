import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { v4 } from 'uuid';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import * as fs from 'fs';
import * as path from 'path';
import { FolderExtensions } from '../../FolderExtensions';
import { CommandInstance } from '../../../../cli';

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
  private readonly fileChunkSize: number = 10 * 1024 * 1024;  // max fileChunkingThreshold
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.contentType = (!(!args.options.contentType)).toString();
    telemetryProps.checkOut = args.options.checkOut || false;
    telemetryProps.checkInComment = (!(!args.options.checkInComment)).toString();
    telemetryProps.approve = args.options.approve || false;
    telemetryProps.approveComment = (!(!args.options.approveComment)).toString();
    telemetryProps.publish = args.options.publish || false;
    telemetryProps.publishComment = (!(!args.options.publishComment)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const folderPath: string = Utils.getServerRelativePath(args.options.webUrl, args.options.folder);
    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = Utils.getSafeFileName(path.basename(fullPath));
    const folderExtensions: FolderExtensions = new FolderExtensions(cmd, this.debug);

    let isCheckedOut: boolean = false;
    let listSettings: ListSettings;

    if (this.debug) {
      cmd.log(`folder path: ${folderPath}...`);
    }

    if (this.debug) {
      cmd.log('Check if the specified folder exists.')
      cmd.log('');
    }

    if (this.debug) {
      cmd.log(`file name: ${fileName}...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      }
    };

    request.get<void>(requestOptions).catch((err: string): Promise<void> => {
      // folder does not exist so will attempt to create the folder tree
      return folderExtensions.ensureFolder(args.options.webUrl, folderPath);
    })
      .then((): Promise<void> => {
        if (args.options.checkOut) {
          return this.fileCheckOut(fileName, args.options.webUrl, folderPath)
            .then((res: any) => {
              // flag the file is checkedOut by the command 
              // so in case of command failure we can try check it in
              isCheckedOut = true;

              return Promise.resolve();
            })
            .catch((err: any) => {
              return Promise.reject(err);
            });
        }

        return Promise.resolve();
      })
      .then((): Promise<void> => {
        if (this.verbose) {
          cmd.log(`Upload file to site ${args.options.webUrl}...`);
        }

        const fileStats: fs.Stats = fs.statSync(fullPath);
        const fileSize: number = fileStats.size;
        if (this.debug) {
          cmd.log(`File size is ${fileSize} bytes`);
        }

        // only up to 250 MB are allowed in a single request
        if (fileSize > this.fileChunkingThreshold) {
          const fileChunkCount: number = Math.ceil(fileSize / this.fileChunkSize);
          if (this.verbose) {
            cmd.log(`Uploading ${fileSize} bytes in ${fileChunkCount} chunks...`);
          }

          // initiate chunked upload session
          const uploadId: string = v4();
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files/GetByPathOrAddStub(DecodedUrl='${encodeURIComponent(fileName)}')/StartUpload(uploadId=guid'${uploadId}')`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            }
          };

          return request
            .post<void>(requestOptions)
            .then((): Promise<void> => {
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

              return new Promise<void>((resolve: () => void, reject: (err: any) => void): void => {
                this.uploadFileChunks(fileUploadInfo, cmd, resolve, reject);
              })
                .then((): Promise<void> => {
                  if (this.verbose) {
                    cmd.log(`Finished uploading ${fileUploadInfo.Position} bytes in ${fileChunkCount} chunks`)
                  }
                  return Promise.resolve();
                })
                .catch((err: any) => {
                  if (this.verbose) {
                    cmd.log('Cancelling upload session due to error...')
                  }

                  const requestOptions: any = {
                    url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files('${encodeURIComponent(fileName)}')/cancelupload(uploadId=guid'${uploadId}')`,
                    headers: {
                      'accept': 'application/json;odata=nometadata'
                    }
                  };

                  return request
                    .post<void>(requestOptions)
                    .then((): Promise<void> => {
                      return Promise.reject(err);  // original error
                    })
                    .catch((err_: any) => {
                      if (this.debug) {
                        cmd.log(`Failed to cancel upload session: ${err_}`);
                      }
                      return Promise.reject(err);  // original error
                    });
                });
            });
        }

        // upload small file in a single request
        const fileBody: Buffer = fs.readFileSync(fullPath);
        const bodyLength: number = fileBody.byteLength;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files/Add(url='${encodeURIComponent(fileName)}', overwrite=true)`,
          body: fileBody,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'content-length': bodyLength
          }
        };

        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
        if (args.options.contentType || args.options.publish || args.options.approve) {
          return this.getFileParentList(fileName, args.options.webUrl, folderPath, cmd)
            .then((listSettingsResp: ListSettings) => {
              listSettings = listSettingsResp;

              if (args.options.contentType) {
                return this.listHasContentType(args.options.contentType, args.options.webUrl, listSettings, cmd);
              }

              return Promise.resolve();
            });
        }

        return Promise.resolve();
      })
      .then((): Promise<void> => {
        // check if there are unknown options 
        // and map them as fields to update
        let fieldsToUpdate: FieldValue[] = this.mapUnknownOptionsAsFieldValue(args.options);

        if (args.options.contentType) {
          fieldsToUpdate.push({
            FieldName: 'ContentType',
            FieldValue: args.options.contentType
          });
        }

        if (fieldsToUpdate.length > 0) {
          // perform list item update and checkin
          return this.validateUpdateListItem(args.options.webUrl, folderPath, fileName, fieldsToUpdate, cmd, args.options.checkInComment);
        }
        else if (isCheckedOut) {
          // perform checkin
          return this.fileCheckIn(args, fileName);
        }

        return Promise.resolve();
      })
      .then((): Promise<void> => {
        // approve and publish cannot be used together 
        // when approve is used it will automatically publish the file
        // so then no need to publish afterwards
        if (args.options.approve) {
          if (this.verbose) {
            cmd.log(`Approve file ${fileName}`);
          }

          // approve the existing file with given comment
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files('${encodeURIComponent(fileName)}')/approve(comment='${encodeURIComponent(args.options.approveComment || '')}')`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            json: true
          };

          return request.post(requestOptions);
        }
        else if (args.options.publish) {
          if (listSettings.EnableModeration && listSettings.EnableMinorVersions) {
            return Promise.reject('The file cannot be published without approval. Moderation for this list is enabled. Use the --approve option instead of --publish to approve and publish the file');
          }

          if (this.verbose) {
            cmd.log(`Publish file ${fileName}`);
          }

          // publish the existing file with given comment
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files('${encodeURIComponent(fileName)}')/publish(comment='${encodeURIComponent(args.options.publishComment || '')}')`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            json: true
          };

          return request.post(requestOptions);
        }

        return Promise.resolve();
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log('DONE');
        }

        cb();
      }, (err: any): void => {
        if (isCheckedOut) {
          // in a case the command has done checkout
          // then have to rollback the checkout

          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files('${encodeURIComponent(fileName)}')/UndoCheckOut()`,
          };

          request.post(requestOptions)
            .then(_ => this.handleRejectedODataJsonPromise(err, cmd, cb))
            .catch(checkoutError => {
              if (this.verbose) {
                cmd.log('Could not rollback file checkout');
                cmd.log(checkoutError);
                cmd.log('');
              }

              this.handleRejectedODataJsonPromise(err, cmd, cb);
            });
        }
        else {
          this.handleRejectedODataJsonPromise(err, cmd, cb);
        }
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the file should be uploaded to'
      },
      {
        option: '-f, --folder <folder>',
        description: 'Site-relative or server-relative URL to the folder where the file should be uploaded'
      },
      {
        option: '-p, --path <path>',
        description: 'Local path to the file to upload'
      },
      {
        option: '-c, --contentType [contentType]',
        description: 'Content type name or ID to assign to the file'
      },
      {
        option: '--checkOut',
        description: 'If versioning is enabled, this will check out the file first if it exists, upload the file, then check it in again'
      },
      {
        option: '--checkInComment [checkInComment]',
        description: 'Comment to set when checking the file in'
      },
      {
        option: '--approve',
        description: 'Will automatically approve the uploaded file'
      },
      {
        option: '--approveComment [approveComment]',
        description: 'Comment to set when approving the file'
      },
      {
        option: '--publish',
        description: 'Will automatically publish the uploaded file'
      },
      {
        option: '--publishComment [publishComment]',
        description: 'Comment to set when publishing the file'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
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
    };
  }

  private listHasContentType(contentType: string, webUrl: string, listSettings: ListSettings, cmd: any): Promise<void> {
    if (this.verbose) {
      cmd.log(`Getting list of available content types ...`);
    }

    const requestOptions: any = {
      url: `${webUrl}/_api/web/lists('${listSettings.Id}')/contenttypes?$select=Name,Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    return request.get<any>(requestOptions).then(response => {
      // check if the specified content type is in the list
      for (const ct of response.value) {
        if (ct.Id.StringValue === contentType || ct.Name === contentType) {
          return Promise.resolve();
        }
      }

      return Promise.reject(`Specified content type '${contentType}' doesn't exist on the target list`);
    });
  }

  private fileCheckOut(fileName: string, webUrl: string, folder: string): Promise<void> {
    // check if file already exists, otherwise it can't be checked out
    const requestOptions: any = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folder)}')/Files('${encodeURIComponent(fileName)}')`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      }
    };

    return request.get<void>(requestOptions)
      .then(() => {
        // checkout the existing file
        const requestOptions: any = {
          url: `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folder)}')/Files('${encodeURIComponent(fileName)}')/CheckOut()`,
          headers: {
            'accept': 'application/json;odata=nometadata',
          },
          json: true
        };

        return request.post<void>(requestOptions);
      });
  }

  private uploadFileChunks(info: FileUploadInfo, cmd: any, resolve: () => void, reject: (err: any) => void): void {
    let fd: number = 0;
    try {
      fd = fs.openSync(info.FilePath, 'r');
      const fileBuffer: Buffer = Buffer.alloc(this.fileChunkSize);
      const readCount: number = fs.readSync(fd, fileBuffer, 0, this.fileChunkSize, info.Position);
      fs.closeSync(fd);
      fd = 0;

      const offset: number = info.Position;
      info.Position += readCount;
      const isLastChunk: boolean = info.Position >= info.Size;

      const requestOptions: any = {
        url: `${info.WebUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(info.FolderPath)}')/Files('${encodeURIComponent(info.Name)}')/${isLastChunk ? 'Finish' : 'Continue'}Upload(uploadId=guid'${info.Id}',fileOffset=${offset})`,
        body: fileBuffer,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-length': readCount
        }
      };

      request
        .post<void>(requestOptions)
        .then((): void => {
          if (this.verbose) {
            cmd.log(`Uploaded ${info.Position} of ${info.Size} bytes (${Math.round(100 * info.Position / info.Size)}%)`);
          }

          if (isLastChunk) {
            resolve();
          }
          else {
            this.uploadFileChunks(info, cmd, resolve, reject);
          }
        })
        .catch((err: any) => {
          if (--info.RetriesLeft > 0) {
            if (this.verbose) {
              cmd.log(`Retrying to upload chunk due to error: ${err}`);
            }
            info.Position -= readCount;  // rewind
            this.uploadFileChunks(info, cmd, resolve, reject);
          }
          else {
            reject(err);
          }
        });
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
          cmd.log(`Retrying to read chunk due to error: ${err}`);
        }
        this.uploadFileChunks(info, cmd, resolve, reject);
      }
      else {
        reject(err);
      }
    }
  }

  private getFileParentList(fileName: string, webUrl: string, folder: string, cmd: any): Promise<ListSettings> {
    if (this.verbose) {
      cmd.log(`Getting list details in order to get its available content types afterwards...`);
    }

    const requestOptions: any = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folder)}')/Files('${encodeURIComponent(fileName)}')/ListItemAllFields/ParentList?$Select=Id,EnableModeration,EnableVersioning,EnableMinorVersions`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    return request.get(requestOptions);
  }

  private validateUpdateListItem(webUrl: string, folderPath: string, fileName: string, fieldsToUpdate: FieldValue[], cmd: any, checkInComment?: string): Promise<void> {
    if (this.verbose) {
      cmd.log(`Validate and update list item values for file ${fileName}`);
    }

    const requestBody: any = {
      formValues: fieldsToUpdate,
      bNewDocumentUpdate: true, // true = will automatically checkin the item, but we will use it to perform system update and also do a checkin
      checkInComment: checkInComment || ''
    };

    if (this.debug) {
      cmd.log('ValidateUpdateListItem will perform the checkin ...');
      cmd.log('');
    }

    // update the existing file list item fields
    const requestOptions: any = {
      url: `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Files('${encodeURIComponent(fileName)}')/ListItemAllFields/ValidateUpdateListItem()`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      body: requestBody,
      json: true
    };

    return request.post(requestOptions)
      .then((res: any) => {
        // check for field value update for errors
        const fieldValues: FieldValueResult[] = res.value;
        for (const fieldValue of fieldValues) {
          if (fieldValue.HasException) {
            return Promise.reject(`Update field value error: ${JSON.stringify(fieldValues)}`);
          }
        }
        return Promise.resolve();
      })
      .catch((err: any) => {
        return Promise.reject(err);
      });
  }

  private fileCheckIn(args: any, fileName: string): Promise<void> {
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files('${encodeURIComponent(fileName)}')/CheckIn(comment='${encodeURIComponent(args.options.checkInComment || '')}',checkintype=0)`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
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
      'output'
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