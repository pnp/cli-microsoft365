import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import AadCommand from '../../../base/AadCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving service principal information...`);
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
      responseType: 'json'
    };

    request
      .get<{ value: any[] }>(requestOptions)
      .then((res: { value: any[] }): void => {
        if (res.value && res.value.length > 0) {
          logger.log(res.value[0]);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --appId [appId]'
      },
      {
        option: '-n, --displayName [displayName]'
      },
      {
        option: '--objectId [objectId]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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
  }
}

module.exports = new AadSpGetCommand();