import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError,
  CommandTypes
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { FieldLink } from './FieldLink';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  contentTypeId: string;
  fieldId: string;
  hidden?: string;
  required?: string;
  webUrl: string;
}

class SpoContentTypeFieldSetCommand extends SpoCommand {
  private requestDigest: string;
  private siteId: string;
  private webId: string;
  private fieldLink: FieldLink | null;

  constructor() {
    super();
    this.requestDigest = '';
    this.siteId = '';
    this.webId = '';
    this.fieldLink = null;
  }

  public get name(): string {
    return `${commands.CONTENTTYPE_FIELD_SET}`;
  }

  public get description(): string {
    return 'Adds or updates a site column reference in a site content type';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.hidden = args.options.hidden;
    telemetryProps.required = args.options.required;
    return telemetryProps;
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['contentTypeId', 'c']
    };
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let schemaXmlWithResourceTokens: string = '';

    if (this.verbose) {
      cmd.log(`Retrieving field link for field ${args.options.fieldId}...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/contenttypes('${encodeURIComponent(args.options.contentTypeId)}')/fieldlinks('${args.options.fieldId}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    }

    request
      .get<FieldLink>(requestOptions)
      .then((res: FieldLink): Promise<{ SchemaXmlWithResourceTokens: string; }> => {
        if (res["odata.null"] !== true) {
          if (this.verbose) {
            cmd.log('Field link found');
          }
          this.fieldLink = res;
          return Promise.resolve(undefined as any);
        }

        if (this.verbose) {
          cmd.log('Field link not found. Creating...');
          cmd.log(`Retrieving information about site column ${args.options.fieldId}...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/fields('${args.options.fieldId}')?$select=SchemaXmlWithResourceTokens`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        }

        return request.get(requestOptions);
      })
      .then((res?: { SchemaXmlWithResourceTokens: string; }): Promise<void> => {
        if (!res) {
          return Promise.resolve();
        }

        schemaXmlWithResourceTokens = res.SchemaXmlWithResourceTokens;
        return this.createFieldLink(cmd, args, schemaXmlWithResourceTokens);
      })
      .then((): Promise<FieldLink> => {
        if (this.fieldLink) {
          return Promise.resolve(undefined as any);
        }

        if (this.verbose) {
          cmd.log(`Retrieving information about field link for field ${args.options.fieldId}...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/contenttypes('${encodeURIComponent(args.options.contentTypeId)}')/fieldlinks('${args.options.fieldId}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((res?: FieldLink): Promise<{ Id: string; }> => {
        if (res && res["odata.null"] !== true) {
          this.fieldLink = res;
        }

        if (!this.fieldLink) {
          return Promise.reject(`Couldn't find field link for field ${args.options.fieldId}`);
        }

        let updateHidden: boolean = false;
        let updateRequired: boolean = false;
        if (typeof args.options.hidden !== 'undefined' &&
          this.fieldLink.Hidden !== (args.options.hidden === 'true')) {
          updateHidden = true;
        }
        if (typeof args.options.required !== 'undefined' &&
          this.fieldLink.Required !== (args.options.required === 'true')) {
          updateRequired = true;
        }

        if (!updateHidden && !updateRequired) {
          if (this.verbose) {
            cmd.log('Field link already up-to-date');
          }
          return Promise.reject('DONE');
        }

        if (this.siteId) {
          return Promise.resolve(undefined as any);
        }

        if (this.verbose) {
          cmd.log(`Retrieving site collection id...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/site?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        }

        return request.get(requestOptions);
      })
      .then((res?: { Id: string }): Promise<{ Id: string; }> => {
        if (res) {
          this.siteId = res.Id;
        }

        if (this.webId) {
          return Promise.resolve(undefined as any);
        }

        if (this.verbose) {
          cmd.log(`Retrieving site id...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        }

        return request.get(requestOptions);
      })
      .then((res?: { Id: string }): Promise<string> => {
        if (res) {
          this.webId = res.Id;
        }

        if (this.verbose) {
          cmd.log(`Updating field link...`);
        }

        const requiredProperty: string = typeof args.options.required !== 'undefined' &&
          (this.fieldLink as FieldLink).Required !== (args.options.required === 'true') ? `<SetProperty Id="122" ObjectPathId="121" Name="Required"><Parameter Type="Boolean">${args.options.required}</Parameter></SetProperty>` : '';
        const hiddenProperty: string = typeof args.options.hidden !== 'undefined' &&
          (this.fieldLink as FieldLink).Hidden !== (args.options.hidden === 'true') ? `<SetProperty Id="123" ObjectPathId="121" Name="Hidden"><Parameter Type="Boolean">${args.options.hidden}</Parameter></SetProperty>` : '';

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': this.requestDigest
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${requiredProperty}${hiddenProperty}<Method Name="Update" Id="124" ObjectPathId="19"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="121" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${Utils.escapeXml(args.options.contentTypeId)}:fl:${(this.fieldLink as FieldLink).Id}" /><Identity Id="19" Name="716a7b9e-3012-0000-22fb-84acfcc67d04|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${Utils.escapeXml(args.options.contentTypeId)}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
        }
        else {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
          cb();
        }
      }, (error: any): void => {
        if (error === 'DONE') {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }
          cb();
        }
        else {
          this.handleRejectedODataJsonPromise(error, cmd, cb);
        }
      });
  }

  private createFieldLink(cmd: CommandInstance, args: CommandArgs, schemaXmlWithResourceTokens: string): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
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

      this
        .updateField(xField, requiresUpdate, cmd, args)
        .then((): Promise<{ Id: string; }> => {
          if (this.verbose) {
            cmd.log(`Retrieving site collection id...`);
          }

          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/site?$select=Id`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            json: true
          }

          return request.get(requestOptions);
        })
        .then((res: { Id: string }): Promise<{ Id: string; }> => {
          this.siteId = res.Id;

          if (this.verbose) {
            cmd.log(`Retrieving site id...`);
          }

          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web?$select=Id`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            json: true
          }

          return request.get(requestOptions);
        })
        .then((res: { Id: string }): Promise<void> => {
          this.webId = res.Id;

          return this.ensureRequestDigest(args.options.webUrl, cmd);
        })
        .then((): Promise<string> => {
          const requestOptions: any = {
            url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': this.requestDigest
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><Method Name="Update" Id="7" ObjectPathId="1"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="2" Name="d6667b9e-50fb-0000-2693-032ae7a0df25|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:field:${args.options.fieldId}" /><Method Id="4" ParentId="3" Name="Add"><Parameters><Parameter TypeId="{63fb2c92-8f65-4bbb-a658-b6cd294403f4}"><Property Name="Field" ObjectPathId="2" /></Parameter></Parameters></Method><Identity Id="1" Name="d6667b9e-80f4-0000-2693-05528ff416bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${this.siteId}:web:${this.webId}:contenttype:${Utils.escapeXml(args.options.contentTypeId)}" /><Property Id="3" ParentId="1" Name="FieldLinks" /></ObjectPaths></Request>`
          };

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            reject(response.ErrorInfo.ErrorMessage);
          }
          else {
            resolve();
          }
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private updateField(schemaXml: string, requiresUpdate: boolean, cmd: CommandInstance, args: CommandArgs): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (!requiresUpdate) {
        if (this.verbose) {
          cmd.log(`Schema of field ${args.options.fieldId} is already up-to-date`);
        }
        resolve();
        return;
      }

      this
        .ensureRequestDigest(args.options.webUrl, cmd)
        .then((): Promise<void> => {
          if (this.verbose) {
            cmd.log(`Updating field schema...`);
          }

          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/fields('${args.options.fieldId}')`,
            headers: {
              accept: 'application/json;odata=nometadata',
              'content-type': 'application/json;odata=nometadata',
              'X-HTTP-Method': 'MERGE',
              'x-requestdigest': this.requestDigest
            },
            body: {
              SchemaXml: schemaXml
            },
            json: true
          }

          return request.post(requestOptions);
        })
        .then((): void => {
          resolve();
        }, (error: any): void => {
          reject(error);
        })
    });
  }

  private ensureRequestDigest(siteUrl: string, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (this.requestDigest) {
        if (this.debug) {
          cmd.log('Request digest already present');
        }
        resolve();
        return;
      }

      if (this.debug) {
        cmd.log('Retrieving request digest...');
      }

      this
        .getRequestDigest(siteUrl)
        .then((res: ContextInfo): void => {
          this.requestDigest = res.FormDigestValue;
          resolve();
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site where the content type is located'
      },
      {
        option: '-c, --contentTypeId <contentTypeId>',
        description: 'ID of the content type on which the field reference should be set'
      },
      {
        option: '-f, --fieldId <fieldId>',
        description: 'ID of the field to which the reference should be set'
      },
      {
        option: '-r, --required [required]',
        description: 'Set to true, if the field should be required or to false if it should be optional'
      },
      {
        option: '--hidden [hidden]',
        description: 'Set to true, if the field should be hidden or to false if it should be visible'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.fieldId)) {
        return `${args.options.fieldId} is not a valid GUID`;
      }

      if (typeof args.options.required !== 'undefined') {
        if (args.options.required !== 'true' &&
          args.options.required !== 'false') {
          return `${args.options.required} is not a valid boolean value. Allowed values are true|false`;
        }
      }

      if (typeof args.options.hidden !== 'undefined') {
        if (args.options.hidden !== 'true' &&
          args.options.hidden !== 'false') {
          return `${args.options.hidden} is not a valid boolean value. Allowed values are true|false`;
        }
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoContentTypeFieldSetCommand();