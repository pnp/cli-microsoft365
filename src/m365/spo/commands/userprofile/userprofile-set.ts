import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ContextInfo } from '../../spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName: string;
  propertyName: string;
  propertyValue: string;
}

class SpoUserProfileSetCommand extends SpoCommand {
  public get name(): string {
    return commands.USERPROFILE_SET;
  }

  public get description(): string {
    return 'Sets user profile property for a SharePoint user';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let spoUrl: string = '';

    this
      .getSpoUrl(logger, this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;

        return this.getRequestDigest(spoUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const propertyValue: string[] = args.options.propertyValue.split(',').map(o => o.trim());
        let propertyType: string = 'SetSingleValueProfileProperty';
        const data: any = {
          accountName: `i:0#.f|membership|${args.options.userName}`,
          propertyName: args.options.propertyName
        };

        if (propertyValue.length > 1) {
          propertyType = 'SetMultiValuedProfileProperty';
          data.propertyValues = [...propertyValue];
        }
        else {
          data.propertyValue = propertyValue[0];
        }

        const requestOptions: any = {
          url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/${propertyType}`,
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'X-RequestDigest': res.FormDigestValue
          },
          data: data,
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --userName <userName>'
      },
      {
        option: '-n, --propertyName <propertyName>'
      },
      {
        option: '-v, --propertyValue <propertyValue>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoUserProfileSetCommand();
