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

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  id?: string;
  classification?: string;
  disableFlows?: string;
}

class SpoSiteSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_SET;
  }

  public get description(): string {
    return 'Updates properties of the specified site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id === 'string';
    telemetryProps.classification = typeof args.options.classification === 'string'
    telemetryProps.disableFlows = typeof args.options.disableFlows === 'string'
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let siteId: string = '';

    ((): Promise<{ Id: string }> => {
      if (this.debug) {
        cmd.log(`Retrieving site ID...`);
      }

      if (args.options.id) {
        return Promise.resolve({ Id: args.options.id });
      }
      else {
        const requestOptions: any = {
          url: `${args.options.url}/_api/site?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      }
    })()
      .then((res: { Id: string }): Promise<ContextInfo> => {
        siteId = res.Id;

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigest(args.options.url);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Updating site ${args.options.url} properties...`);
        }

        const classification: string = typeof args.options.classification === 'string' ? `<SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">${Utils.escapeXml(args.options.classification)}</Parameter></SetProperty>` : '';
        const disableFlows: string = typeof args.options.disableFlows === 'string' ? `<SetProperty Id="28" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">${args.options.disableFlows === 'true'}</Parameter></SetProperty>` : '';

        const requestOptions: any = {
          url: `${args.options.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${classification}${disableFlows}</Actions><ObjectPaths><Identity Id="5" Name="e10a459e-60c8-4000-8240-a68d6a12d39e|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}" /></ObjectPaths></Request>`
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
        else {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The URL of the site collection to update'
      },
      {
        option: '-i, --id [id]',
        description: 'The ID of the site collection to update'
      },
      {
        option: '--classification [classification]',
        description: 'The new classification for the site collection'
      },
      {
        option: '--disableFlows [disableFlows]',
        description: 'Set to true to disable using Microsoft Flow in this site collection'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.url) {
        return 'Required parameter url missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.id &&
        !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (typeof args.options.classification === 'undefined' &&
        typeof args.options.disableFlows === 'undefined') {
        return 'Specify at least one property to update';
      }

      if (typeof args.options.disableFlows === 'string' &&
        args.options.disableFlows !== 'true' &&
        args.options.disableFlows !== 'false') {
        return `${args.options.disableFlows} is not a valid value for the disableFlow option. Allowed values are true|false`;
      }

      return true;
    };
  }

  public types(): CommandTypes {
    // required to support passing empty strings as valid values
    return {
      string: ['classification']
    }
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.SITE_SET).helpInformation());
    log(
      `  Remarks:

    If the specified ${chalk.grey('url')} doesn't refer to an existing site collection,
    you will get a ${chalk.grey('404 - "404 FILE NOT FOUND"')} error.

    To update site collection's properties, the command requires site collection
    ID. The command can retrieve it automatically, but if you already have it,
    you can save an additional request, by specifying it using the ${chalk.grey('id')}
    option.

  Examples:
  
    Update site collection's classification. Will automatically retrieve the ID
    of the site collection
      ${this.name} --url https://contoso.sharepoint.com/sites/sales --classification MBI

    Reset site collection's classification.
      ${this.name} --url https://contoso.sharepoint.com/sites/sales --id 255a50b2-527f-4413-8485-57f4c17a24d1 --classification

    Disable using Microsoft Flow on the site collection. Will automatically retrieve the
    ID of the site collection
      ${this.name} --url https://contoso.sharepoint.com/sites/sales --disableFlows true
`);
  }
}

module.exports = new SpoSiteSetCommand();