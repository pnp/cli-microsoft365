import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import AadCommand from '../../../base/AadCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  displayName?: string;
  objectId?: string;
}

class AadSpGetCommand extends AadCommand {
  public get name(): string {
    return commands.SP_GET;
  }

  public get description(): string {
    return 'Gets information about the specific service principal';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = (!(!args.options.appId)).toString();
    telemetryProps.displayName = (!(!args.options.displayName)).toString();
    telemetryProps.objectId = (!(!args.options.objectId)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving service principal information...`);
    }

    let spMatchQuery: string = '';
    if (args.options.appId) {
      spMatchQuery = `appId eq '${encodeURIComponent(args.options.appId)}'`;
    }
    else if (args.options.objectId) {
      spMatchQuery = `objectId eq '${encodeURIComponent(args.options.objectId)}'`;
    }
    else {
      spMatchQuery = `displayName eq '${encodeURIComponent(args.options.displayName as string)}'`;
    }

    const requestOptions: any = {
      url: `${this.resource}/myorganization/servicePrincipals?api-version=1.6&$filter=${spMatchQuery}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<{ value: any[] }>(requestOptions)
      .then((res: { value: any[] }): void => {
        if (res.value && res.value.length > 0) {
          cmd.log(res.value[0]);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --appId [appId]',
        description: 'ID of the application for which the service principal should be retrieved'
      },
      {
        option: '-n, --displayName [displayName]',
        description: 'Display name of the application for which the service principal should be retrieved'
      },
      {
        option: '--objectId [objectId]',
        description: 'ObjectId of the application for which the service principal should be retrieved'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      let optionsSpecified: number = 0;
      optionsSpecified += args.options.appId ? 1 : 0;
      optionsSpecified += args.options.displayName ? 1 : 0;
      optionsSpecified += args.options.objectId ? 1 : 0;
      if (optionsSpecified !== 1) {
        return 'Specify either appId, objectId or displayName';
      }

      if (args.options.appId && !Utils.isValidGuid(args.options.appId)) {
        return `${args.options.appId} is not a valid appId GUID`;
      }

      if (args.options.objectId && !Utils.isValidGuid(args.options.objectId)) {
        return `${args.options.objectId} is not a valid objectId GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.SP_GET).helpInformation());
    log(
      `  Remarks:
  
    Specify either the ${chalk.grey('appId')}, ${chalk.grey('objectId')} or ${chalk.grey('displayName')}. 
    If you specify more than one option value, the command will fail
    with an error.
   
  Examples:
  
    Return details about the service principal with appId ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}.
      ${commands.SP_GET} --appId b2307a39-e878-458b-bc90-03bc578531d6

    Return details about the ${chalk.grey('Microsoft Graph')} service principal.
      ${commands.SP_GET} --displayName "Microsoft Graph"

    Return details about the service principal with ObjectId ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531dd')}.
      ${commands.SP_GET} --objectId b2307a39-e878-458b-bc90-03bc578531dd

  More information:
  
    Application and service principal objects in Azure Active Directory (Azure AD)
      https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects
`);
  }
}

module.exports = new AadSpGetCommand();