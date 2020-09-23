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
        logger.log(this.formatOutput(propertyBagData));

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
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

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
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