import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import { validation } from '../../../../utils';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  name: string;
  description?: string;
  allowMembersEditMembership?: boolean;
  onlyAllowMembersViewMembership?: boolean;
  allowRequestToJoinLeave?: boolean;
  autoAcceptRequestToJoinLeave?: boolean;
  requestToJoinLeaveEmailSetting?: string;
}

class SpoGroupAddCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_ADD;
  }

  public get description(): string {
    return 'Creates group in the specified site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.allowMembersEditMembership = typeof args.options.allowMembersEditMembership !== 'undefined';
    telemetryProps.onlyAllowMembersViewMembership = typeof args.options.onlyAllowMembersViewMembership !== 'undefined';
    telemetryProps.allowRequestToJoinLeave = typeof args.options.allowRequestToJoinLeave !== 'undefined';
    telemetryProps.autoAcceptRequestToJoinLeave = typeof args.options.autoAcceptRequestToJoinLeave !== 'undefined';
    telemetryProps.requestToJoinLeaveEmailSetting = typeof args.options.requestToJoinLeaveEmailSetting !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: AxiosRequestConfig = {
      url: `${args.options.webUrl}/_api/web/sitegroups`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        Title: args.options.name,
        Description: args.options.description,
        AllowMembersEditMembership: args.options.allowMembersEditMembership,
        OnlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
        AllowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
        AutoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
        RequestToJoinLeaveEmailSetting: args.options.requestToJoinLeaveEmailSetting
      }
    };

    request
      .post(requestOptions)
      .then((response: any): void => {
        logger.log(response);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '--description [description]'
      },
      {
        option: '--allowMembersEditMembership [allowMembersEditMembership]'
      },
      {
        option: '--onlyAllowMembersViewMembership [onlyAllowMembersViewMembership]'
      },
      {
        option: '--allowRequestToJoinLeave [allowRequestToJoinLeave]'
      },
      {
        option: '--autoAcceptRequestToJoinLeave [autoAcceptRequestToJoinLeave]'
      },
      {
        option: '--requestToJoinLeaveEmailSetting [requestToJoinLeaveEmailSetting]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    const booleanOptions = [
      args.options.allowMembersEditMembership, args.options.onlyAllowMembersViewMembership,
      args.options.allowRequestToJoinLeave, args.options.autoAcceptRequestToJoinLeave
    ];
    for (const option of booleanOptions) {
      if (typeof option !== 'undefined' && !validation.isValidBoolean(option as any)) {
        return `Value '${option}' is not a valid boolean`;
      }
    }

    return true;
  }
}

module.exports = new SpoGroupAddCommand();