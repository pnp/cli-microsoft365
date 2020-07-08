import commands from '../../commands';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import { SpoPropertyBagBaseCommand, Property } from './propertybag-base';
import GlobalOptions from '../../../../GlobalOptions';
import { ClientSvc, IdentityResponse } from '../../ClientSvc';

const vorpal: Vorpal = require('../../../../vorpal-init');

export interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  folder?: string;
}

class SpoPropertyBagListCommand extends SpoPropertyBagBaseCommand {

  public get name(): string {
    return `${commands.PROPERTYBAG_LIST}`;
  }

  public get description(): string {
    return 'Gets property bag values';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.folder = (!(!args.options.folder)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const clientSvcCommons: ClientSvc = new ClientSvc(cmd, this.debug);

    this
      .getRequestDigest(args.options.webUrl)
      .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
        this.formDigestValue = contextResponse.FormDigestValue;

        return clientSvcCommons.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
      })
      .then((identityResp: IdentityResponse): Promise<any> => {
        const opts: Options = args.options;
        if (opts.folder) {
          return this.getFolderPropertyBag(identityResp, opts.webUrl, opts.folder, cmd);
        }

        return this.getWebPropertyBag(identityResp, opts.webUrl, cmd);
      })
      .then((propertyBagData: any): void => {
        cmd.log(this.formatOutput(propertyBagData));

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site from which the property bag value should be retrieved'
      },
      {
        option: '-f, --folder [folder]',
        description: 'Site-relative URL of the folder from which to retrieve property bag value. Case-sensitive',
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (SpoCommand.isValidSharePointUrl(args.options.webUrl) !== true) {
        return 'Missing required option url';
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.PROPERTYBAG_LIST).helpInformation());
    log(
      `  Examples:

    Return property bag values located in site ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_LIST} --webUrl https://contoso.sharepoint.com/sites/test

    Return property bag values located in site root folder ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_LIST} --webUrl https://contoso.sharepoint.com/sites/test -f /

    Return property bag values located in site document library ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_LIST} --webUrl https://contoso.sharepoint.com/sites/test --folder '/Shared Documents'

    Return property bag values located in folder in site document library ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_LIST} -w https://contoso.sharepoint.com/sites/test -f '/Shared Documents/MyFolder'

    Return property bag values located in site list ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_LIST} --webUrl https://contoso.sharepoint.com/sites/test --folder /Lists/MyList
    `);
  }

  /**
   * The property bag data returned from the client.svc/ProcessQuery response
   * has to be formatted before displayed since the key, value objects
   * carry extra information.
   * @param propertyBag client.svc property bag javascript object
   */
  private formatOutput(propertyBag: any): Property[] {
    let result: Property[] = [];
    const keys = Object.keys(propertyBag);

    for (let i = 0; i < keys.length; i++) {

      if (keys[i] === '_ObjectType_') {
        // this is system data, do not include it
        continue;
      }
      const formattedProp = this.formatProperty(keys[i], propertyBag[keys[i]]);
      result.push(formattedProp);
    }
    return result;
  }
}

module.exports = new SpoPropertyBagListCommand();