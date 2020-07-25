import { ContextInfo } from '../../spo';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoUrl: string = '';

    this
      .getSpoUrl(cmd, this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;

        return this.getRequestDigest(spoUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const propertyValue: string[] = args.options.propertyValue.split(',').map(o => o.trim());
        let propertyType: string = 'SetSingleValueProfileProperty';
        const body: any = {
          accountName: `i:0#.f|membership|${args.options.userName}`,
          propertyName: args.options.propertyName
        };

        if (propertyValue.length > 1) {
          propertyType = 'SetMultiValuedProfileProperty';
          body.propertyValues = [...propertyValue];
        }
        else {
          body.propertyValue = propertyValue[0];
        }

        const requestOptions: any = {
          url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/${propertyType}`,
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'X-RequestDigest': res.FormDigestValue
          },
          body: body,
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.debug) {
          cmd.log(chalk.green('DONE'));
        }
        
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --userName <userName>',
        description: 'Account name of the user'
      },
      {
        option: '-n, --propertyName <propertyName>',
        description: 'The name of the property to be set'
      },
      {
        option: '-v, --propertyValue <propertyValue>',
        description: 'The value of the property to be set'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoUserProfileSetCommand();
