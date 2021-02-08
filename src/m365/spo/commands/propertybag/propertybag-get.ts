import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import SpoCommand from '../../../base/SpoCommand';
import { ClientSvc, IdentityResponse } from '../../ClientSvc';
import commands from '../../commands';
import { ContextInfo } from '../../spo';
import { Property, SpoPropertyBagBaseCommand } from './propertybag-base';

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
    return commands.PROPERTYBAG_GET;
  }

  public get description(): string {
    return 'Gets the value of the specified property from the property bag';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.folder = (!(!args.options.folder)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const clientSvcCommons: ClientSvc = new ClientSvc(logger, this.debug);

    this
      .getRequestDigest(args.options.webUrl)
      .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
        this.formDigestValue = contextResponse.FormDigestValue;

        return clientSvcCommons.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
      })
      .then((identityResp: IdentityResponse): Promise<any> => {
        const opts: Options = args.options;
        if (opts.folder) {
          return this.getFolderPropertyBag(identityResp, opts.webUrl, opts.folder, logger);
        }

        return this.getWebPropertyBag(identityResp, opts.webUrl, logger);
      })
      .then((propertyBagData: any): void => {
        const property = this.filterByKey(propertyBagData, args.options.key);

        if (property) {
          logger.log(property.value);
        } else if (this.verbose) {
          logger.logToStderr('Property not found.');
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-k, --key <key>'
      },
      {
        option: '-f, --folder [folder]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
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