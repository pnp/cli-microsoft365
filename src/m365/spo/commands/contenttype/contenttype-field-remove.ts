import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentTypeId: string;
  id: string;
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  updateChildContentTypes?: boolean;
  force?: boolean;
}

class SpoContentTypeFieldRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_FIELD_REMOVE;
  }

  public get description(): string {
    return 'Removes a column from a site- or list content type';
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
        listTitle: typeof args.options.listTitle !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        updateChildContentTypes: !!args.options.updateChildContentTypes,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listTitle [listTitle]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--contentTypeId <contentTypeId>'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-c, --updateChildContentTypes'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('i', 'contentTypeId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeFieldLink = async (): Promise<void> => {
      try {
        if (this.debug) {
          await logger.logToStderr(`Get SiteId required by ProcessQuery endpoint.`);
        }

        const siteId = await spo.getSiteIdBySPApi(args.options.webUrl, logger, this.verbose);

        if (this.debug) {
          await logger.logToStderr(`SiteId: ${siteId}`);
          await logger.logToStderr(`Get WebId required by ProcessQuery endpoint.`);
        }

        const webId = await spo.getWebId(args.options.webUrl, logger, this.verbose);

        if (this.debug) {
          await logger.logToStderr(`WebId: ${webId}`);
        }

        let listId: string | undefined = undefined;

        if (args.options.listId) {
          listId = args.options.listId;
        }

        if (args.options.listTitle || args.options.listUrl) {
          listId = await spo.getListId(args.options.webUrl, args.options.listTitle, args.options.listUrl, logger, this.verbose);
        }

        if (this.debug) {
          await logger.logToStderr(`ListId: ${listId}`);
        }

        const reqDigest = await spo.getRequestDigest(args.options.webUrl);
        const requestDigest: string = reqDigest.FormDigestValue;

        const updateChildContentTypes: boolean = args.options.listTitle || args.options.listId || args.options.listUrl ? false : args.options.updateChildContentTypes === true;

        if (this.debug) {
          const additionalLog = args.options.listTitle ? `; ListTitle='${args.options.listTitle}'` : args.options.listId ? `; ListId='${args.options.listId}'` : args.options.listUrl ? `; ListUrl='${args.options.listUrl}'` : ` ; UpdateChildContentTypes='${updateChildContentTypes}`;
          await logger.logToStderr(`Remove FieldLink from ContentType. Id='${args.options.id}' ; ContentTypeId='${args.options.contentTypeId}' ${additionalLog}`);
          await logger.logToStderr(`Execute ProcessQuery.`);
        }

        let requestBody: string = '';
        if (listId) {
          requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><Method Name="DeleteObject" Id="21" ObjectPathId="19" /><Method Name="Update" Id="22" ObjectPathId="15"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="17" ParentId="15" Name="FieldLinks" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.id)}}</Parameter></Parameters></Method><Identity Id="15" Name="09eec89e-709b-0000-558c-c222dcaf9162|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:list:${listId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`;
        }
        else {
          requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="77" ObjectPathId="76" /><ObjectPath Id="79" ObjectPathId="78" /><Method Name="DeleteObject" Id="80" ObjectPathId="78" /><Method Name="Update" Id="81" ObjectPathId="24"><Parameters><Parameter Type="Boolean">${updateChildContentTypes}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="76" ParentId="24" Name="FieldLinks" /><Method Id="78" ParentId="76" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.id)}}</Parameter></Parameters></Method><Identity Id="24" Name="6b3ec69e-00a7-0000-55a3-61f8d779d2b3|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:web:${webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`;
        }

        const requestOptions = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest
          },
          data: requestBody
        };

        const res = await request.post<string>(requestOptions);
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeFieldLink();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the column ${args.options.id} from content type ${args.options.contentTypeId}?` });

      if (result) {
        await removeFieldLink();
      }
    }
  }
}

export default new SpoContentTypeFieldRemoveCommand();