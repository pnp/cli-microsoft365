import { ContextInfo } from '../../spo';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName: string;
  propertyName: string;
  propertyValue: string;
}

class SpoUserprofileSetCommand extends SpoCommand {
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
        const propertyValue: string | string[] = args.options.propertyValue.split(',').map(o => o.trim());
        let propertyType: string = 'SetSingleValueProfileProperty';
        const body: any = {
          'accountName': `i:0#.f|membership|${args.options.userName}`,
          'propertyName': args.options.propertyName,
        };

        if (propertyValue.length > 1) {
          propertyType = 'SetMultiValuedProfileProperty';
          body.propertyValues = [...propertyValue];
        } else {
          body.propertyValue = propertyValue[0];
        }

        const requestOptions: any = {
          url: `${spoUrl}/_api/SP.UserProfiles.PeopleManager/${propertyType}`,
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'X-RequestDigest': res.FormDigestValue
          },
          body: JSON.stringify(body)
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --userName <userName>',
        description: 'Account name of the user'
      },
      {
        option: '-n, --propertyName <propertyName>',
        description: 'The property name of the property to be set'
      },
      {
        option: '-v, --propertyValue <propertyValue>',
        description: 'The value of the property to be set'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.userName) {
        return 'Required parameter userName missing';
      }

      if (!args.options.propertyName) {
        return 'Required parameter propertyName missing';
      }

      if (!args.options.propertyValue) {
        return 'Required parameter propertyValue missing';
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    This command requires tenant admin permissions in case of updating properties other than the current logged user.

  Examples:
  
    Updates single value property of a user profile with property name ${chalk.grey('AboutMe')} and property value 'Working as a Microsoft 365 developer'
      ${commands.USERPROFILE_SET} --userName 'john.doe@mytenant.onmicrosoft.com' --propertyName 'AboutMe' --propertyValue 'Working as a Microsoft 365 developer'
  
    Updates multi value property of a user profile with property name ${chalk.grey('SPS-Skills')} and property values 'CSS', 'HTML'
      ${commands.USERPROFILE_SET} --userName 'john.doe@mytenant.onmicrosoft.com' --propertyName 'SPS-Skills' --propertyValue 'CSS, HTML'
`);
  }
}

module.exports = new SpoUserprofileSetCommand();