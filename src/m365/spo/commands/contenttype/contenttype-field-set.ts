import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FieldLink } from './FieldLink';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentTypeId: string;
  id: string;
  hidden?: boolean;
  required?: boolean;
  webUrl: string;
}

class SpoContentTypeFieldSetCommand extends SpoCommand {
  private requestDigest: string;
  private siteId: string;
  private webId: string;
  private fieldLink: FieldLink | null;

  public get name(): string {
    return commands.CONTENTTYPE_FIELD_SET;
  }

  public get description(): string {
    return 'Adds or updates a site column reference in a site content type';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();

    this.requestDigest = '';
    this.siteId = '';
    this.webId = '';
    this.fieldLink = null;
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        hidden: args.options.hidden,
        required: args.options.required
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-c, --contentTypeId <contentTypeId>'
      },
      {
        option: '-f, --id <id>'
      },
      {
        option: '-r, --required [required]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--hidden [hidden]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('contentTypeId', 'c');
    this.types.boolean.push('required', 'hidden');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let schemaXmlWithResourceTokens: string = '';

      if (this.verbose) {
        logger.logToStderr(`Retrieving field link for field ${args.options.id}...`);
      }

      let requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/contenttypes('${formatting.encodeQueryParameter(args.options.contentTypeId)}')/fieldlinks('${args.options.id}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const fieldLink = await request.get<FieldLink>(requestOptions);

      if (fieldLink["odata.null"] !== true) {
        if (this.verbose) {
          logger.logToStderr('Field link found');
        }
        this.fieldLink = fieldLink;
      }
      else {
        if (this.verbose) {
          logger.logToStderr('Field link not found. Creating...');
          logger.logToStderr(`Retrieving information about site column ${args.options.id}...`);
        }

        requestOptions = {
          url: `${args.options.webUrl}/_api/web/fields('${args.options.id}')?$select=SchemaXmlWithResourceTokens`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const field = await request.get<{ SchemaXmlWithResourceTokens: string; }>(requestOptions);
        schemaXmlWithResourceTokens = field.SchemaXmlWithResourceTokens;
        await this.createFieldLink(logger, args, schemaXmlWithResourceTokens);
      }
      if (!this.fieldLink) {
        if (this.verbose) {
          logger.logToStderr(`Retrieving information about field link for field ${args.options.id}...`);
        }

        requestOptions = {
          url: `${args.options.webUrl}/_api/web/contenttypes('${formatting.encodeQueryParameter(args.options.contentTypeId)}')/fieldlinks('${args.options.id}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const fieldLinkResult = await request.get<FieldLink>(requestOptions);
        if (fieldLinkResult && fieldLinkResult["odata.null"] !== true) {
          this.fieldLink = fieldLinkResult;
        }
      }

      if (!this.fieldLink) {
        throw `Couldn't find field link for field ${args.options.id}`;
      }

      let updateHidden: boolean = false;
      let updateRequired: boolean = false;
      if (typeof args.options.hidden !== 'undefined' &&
        this.fieldLink.Hidden !== args.options.hidden) {
        updateHidden = true;
      }
      if (typeof args.options.required !== 'undefined' &&
        this.fieldLink.Required !== args.options.required) {
        updateRequired = true;
      }

      if (!updateHidden && !updateRequired) {
        if (this.verbose) {
          logger.logToStderr('Field link already up-to-date');
        }
        throw 'DONE';
      }

      if (!this.siteId) {
        if (this.verbose) {
          logger.logToStderr(`Retrieving site collection id...`);
        }

        requestOptions = {
          url: `${args.options.webUrl}/_api/site?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const site = await request.get<{ Id: string }>(requestOptions);
        this.siteId = site.Id;
      }

      if (!this.webId) {
        if (this.verbose) {
          logger.logToStderr(`Retrieving site id...`);
        }

        requestOptions = {
          url: `${args.options.webUrl}/_api/web?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const web = await request.get<{ Id: string }>(requestOptions);
        this.webId = web.Id;
      }

      if (this.verbose) {
        logger.logToStderr(`Updating field link...`);
      }

      const requiredProperty: string = typeof args.options.required !== 'undefined' &&
        (this.fieldLink as FieldLink).Required !== args.options.required ? `<SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">${args.options.required}</Parameter></SetProperty>` : '';
      const hiddenProperty: string = typeof args.options.hidden !== 'undefined' &&
        (this.fieldLink as FieldLink).Hidden !== args.options.hidden ? `<SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">${args.options.hidden}</Parameter></SetProperty>` : '';

      requestOptions = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': this.requestDigest
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${requiredProperty}${hiddenProperty}<Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}:fl:${(this.fieldLink as FieldLink).Id}" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
    }
    catch (err: any) {
      if (err !== 'DONE') {
        this.handleRejectedODataJsonPromise(err);
      }
    }
  }

  private async createFieldLink(logger: Logger, args: CommandArgs, schemaXmlWithResourceTokens: string): Promise<void> {

    let requiresUpdate: boolean = false;
    const match: RegExpExecArray = /(<Field[^>]+>)(.*)/.exec(schemaXmlWithResourceTokens) as RegExpExecArray;
    let xField: string = match[1];
    const allowDeletion: RegExpExecArray | null = /AllowDeletion="([^"]+)"/.exec(xField);
    if (!allowDeletion) {
      requiresUpdate = true;
      xField = xField.replace('>', ' AllowDeletion="TRUE">') + match[2];
    }
    else {
      if (allowDeletion[1] !== 'TRUE') {
        requiresUpdate = true;
        xField = xField.replace(allowDeletion[0], 'AllowDeletion="TRUE"') + match[2];
      }
    }

    await this.updateField(xField, requiresUpdate, logger, args);

    if (this.verbose) {
      logger.logToStderr(`Retrieving site collection id...`);
    }

    const requestOptionsSiteId: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/site?$select=Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const resSiteId = await request.get<{ Id: string }>(requestOptionsSiteId);
    this.siteId = resSiteId.Id;

    if (this.verbose) {
      logger.logToStderr(`Retrieving site id...`);
    }

    const requestOptionsWebId: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web?$select=Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const resWebId = await request.get<{ Id: string }>(requestOptionsWebId);

    this.webId = resWebId.Id;

    await this.ensureRequestDigest(args.options.webUrl, logger);

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': this.requestDigest
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:field:${args.options.id}" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${formatting.escapeXml(args.options.contentTypeId)}" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`
    };

    const res = await request.post<string>(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];
    if (response.ErrorInfo) {
      throw response.ErrorInfo.ErrorMessage;
    }
    else {
      return;
    }
  }

  private async updateField(schemaXml: string, requiresUpdate: boolean, logger: Logger, args: CommandArgs): Promise<void> {
    if (!requiresUpdate) {
      if (this.verbose) {
        logger.logToStderr(`Schema of field ${args.options.id} is already up-to-date`);
      }
      return;
    }

    await this.ensureRequestDigest(args.options.webUrl, logger);

    if (this.verbose) {
      logger.logToStderr(`Updating field schema...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/fields('${args.options.id}')`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE',
        'x-requestdigest': this.requestDigest
      },
      data: {
        SchemaXml: schemaXml
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private async ensureRequestDigest(siteUrl: string, logger: Logger): Promise<void> {
    if (this.requestDigest) {
      if (this.debug) {
        logger.logToStderr('Request digest already present');
      }
      return;
    }

    if (this.debug) {
      logger.logToStderr('Retrieving request digest...');
    }

    const res: ContextInfo = await spo.getRequestDigest(siteUrl);
    this.requestDigest = res.FormDigestValue;
  }
}

module.exports = new SpoContentTypeFieldSetCommand();