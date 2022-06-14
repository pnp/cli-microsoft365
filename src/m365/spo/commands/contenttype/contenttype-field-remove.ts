import { Cli, Logger } from '../../../../cli';
import { CommandError, CommandOption, CommandTypes } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentTypeId: string;
  fieldLinkId: string;
  webUrl: string;
  listTitle?: string;
  updateChildContentTypes?: boolean;
  confirm?: boolean;
}

class SpoContentTypeFieldRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_FIELD_REMOVE;
  }

  public get description(): string {
    return 'Removes a column from a site- or list content type';
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['i', 'contentTypeId']
    };
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.updateChildContentTypes = (!(!args.options.updateChildContentTypes)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let webId: string = '';
    let siteId: string = '';
    let listId: string = '';

    const removeFieldLink = (): void => {
      if (this.debug) {
        logger.logToStderr(`Get SiteId required by ProcessQuery endpoint.`);
      }

      // GET SiteId
      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/site?$select=Id`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .get<{ Id: string }>(requestOptions)
        .then((res: { Id: string }): Promise<{ Id: string; }> => {
          siteId = res.Id;

          if (this.debug) {
            logger.logToStderr(`SiteId: ${siteId}`);
            logger.logToStderr(`Get WebId required by ProcessQuery endpoint.`);
          }

          // GET WebId
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web?$select=Id`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          return request.get(requestOptions);
        })
        .then((res: { Id: string }): Promise<{ Id: string; }> => {
          webId = res.Id;

          if (this.debug) {
            logger.logToStderr(`WebId: ${webId}`);
          }

          // If ListTitle is provided
          if (!args.options.listTitle) {
            return Promise.resolve(undefined as any);
          }
          // Request for the ListId
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')?$select=Id`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          return request.get(requestOptions);
        })
        .then((res?: { Id: string }): Promise<ContextInfo> => {
          if (res) {
            listId = res.Id;

            if (this.debug) {
              logger.logToStderr(`ListId: ${listId}`);
            }
          }

          return spo.getRequestDigest(args.options.webUrl);
        })
        .then((res: ContextInfo): Promise<string> => {
          const requestDigest: string = res.FormDigestValue;

          const updateChildContentTypes: boolean = args.options.listTitle ? false : args.options.updateChildContentTypes === true;

          if (this.debug) {
            const additionalLog = args.options.listTitle ? `; ListTitle='${args.options.listTitle}'` : ` ; UpdateChildContentTypes='${updateChildContentTypes}`;
            logger.logToStderr(`Remove FieldLink from ContentType. FieldLinkId='${args.options.fieldLinkId}' ; ContentTypeId='${args.options.contentTypeId}' ${additionalLog}`);
            logger.logToStderr(`Execute ProcessQuery.`);
            logger.logToStderr('');
          }

          let requestBody: string = '';
          if (listId) {
            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><Method Name="DeleteObject" Id="21" ObjectPathId="19" /><Method Name="Update" Id="22" ObjectPathId="15"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="17" ParentId="15" Name="FieldLinks" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.fieldLinkId)}}</Parameter></Parameters></Method><Identity Id="15" Name="09eec89e-709b-0000-558c-c222dcaf9162|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`;
          }
          else {
            requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.fieldLinkId)}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`;
          }

          const requestOptions: any = {
            url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': requestDigest
            },
            data: requestBody
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            cb(new CommandError(response.ErrorInfo.ErrorMessage));
            return;
          }

          cb();
        }, (error: any): void => {
          this.handleRejectedODataJsonPromise(error, logger, cb);
        });
    };

    if (args.options.confirm) {
      removeFieldLink();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the column ${args.options.fieldLinkId} from content type ${args.options.contentTypeId}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeFieldLink();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listTitle [listTitle]'
      },
      {
        option: '-i, --contentTypeId <contentTypeId>'
      },
      {
        option: '-f, --fieldLinkId <fieldLinkId>'
      },
      {
        option: '-c, --updateChildContentTypes'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!validation.isValidGuid(args.options.fieldLinkId)) {
      return `${args.options.fieldLinkId} is not a valid GUID`;
    }

    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoContentTypeFieldRemoveCommand();