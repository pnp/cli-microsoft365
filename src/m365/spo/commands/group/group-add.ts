import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import {CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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
    telemetryProps.description = !!args.options.description;
    telemetryProps.allowMembersEditMembership = !!args.options.allowMembersEditMembership;
    telemetryProps.onlyAllowMembersViewMembership = !!args.options.onlyAllowMembersViewMembership;
    telemetryProps.allowRequestToJoinLeave = !!args.options.allowRequestToJoinLeave;
    telemetryProps.autoAcceptRequestToJoinLeave = !!args.options.autoAcceptRequestToJoinLeave;
    telemetryProps.requestToJoinLeaveEmailSetting = !!args.options.requestToJoinLeaveEmailSetting;
    return telemetryProps;
  }

  public async commandAction(logger: Logger, args: CommandArgs, cb: () => void): Promise<void> {
    try {
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

      const response = await request.post(requestOptions);

      logger.log(response);
      cb();
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err, logger, cb);
    }
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
        option: '--allowMembersEditMembership'
      },
      {
        option: '--onlyAllowMembersViewMembership'
      },
      {
        option: '--allowRequestToJoinLeave'
      },
      {
        option: '--autoAcceptRequestToJoinLeave'
      },
      {
        option: '--requestToJoinLeaveEmailSetting [requestToJoinLeaveEmailSetting]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoGroupAddCommand();