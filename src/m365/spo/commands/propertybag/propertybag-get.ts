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
  key: string;
  folder?: string;
}

class SpoPropertyBagGetCommand extends SpoPropertyBagBaseCommand {
  public get name(): string {
    return `${commands.PROPERTYBAG_GET}`;
  }

  public get description(): string {
    return 'Gets the value of the specified property from the property bag';
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
        const property = this.filterByKey(propertyBagData, args.options.key);

        if (property) {
          cmd.log(property.value);
        } else if (this.verbose) {
          cmd.log('Property not found.');
        }
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
        option: '-k, --key <key>',
        description: 'Key of the property for which the value should be retrieved. Case-sensitive'
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
      if (!args.options.key) {
        return `Required option key missing`;
      }

      if (SpoCommand.isValidSharePointUrl(args.options.webUrl) !== true) {
        return 'Missing required option url';
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.PROPERTYBAG_GET).helpInformation());
    log(
      `  Examples:

    Returns the value of the ${chalk.grey('key1')} property from the property bag located
    in site ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_GET} --webUrl https://contoso.sharepoint.com/sites/test --key key1
    
    Returns the value of the ${chalk.grey('key1')} property from the property bag located
    in site root folder ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_GET} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder /

    Returns the value of the ${chalk.grey('key1')} property from the property bag located
    in site document library ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_GET} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder '/Shared Documents'

    Returns the value of the ${chalk.grey('key1')} property from the property bag located
    in folder in site document library ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_GET} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder '/Shared Documents/MyFolder'

    Returns the value of the ${chalk.grey('key1')} property from the property bag located
    in site list ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.PROPERTYBAG_GET} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder /Lists/MyList
      `);
  }

  private filterByKey(propertyBag: any, key: string): Property | null {
    const keys = Object.keys(propertyBag);

    for (let i = 0; i < keys.length; i++) {
      // we have to normalize the keys and values before we can filter
      // since they carry extra information
      // ex. : 'vti_level$  Int32' instead of 'vti_level'
      const formattedProperty = this.formatProperty(keys[i], propertyBag[keys[i]]);
      if (formattedProperty.key === key) {
        return formattedProperty;
      }
    }

    return null;
  }
}

module.exports = new SpoPropertyBagGetCommand();