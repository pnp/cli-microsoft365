import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  displayName?: string;
  description?: string;
  mailNickName?: string;
  classification?: string;
  visibility?: string;
}

class TeamsSetCommand extends GraphCommand {
  private static props: string[] = [
    'displayName',
    'description',
    'mailNickName',
    'classification',
    'visibility '
  ];

  public get name(): string {
    return `${commands.TEAMS_TEAM_SET}`;
  }

  public get description(): string {
    return 'Updates settings of a Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    TeamsSetCommand.props.forEach((p: string) => {
      telemetryProps[p] = typeof (args.options as any)[p] !== 'undefined';
    });
    return telemetryProps;
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};
    if (options.displayName) {
      requestBody.displayName = options.displayName;
    }
    if (options.description) {
      requestBody.description = options.description;
    }
    if (options.mailNickName) {
      requestBody.mailNickName = options.mailNickName;
    }
    if (options.classification) {
      requestBody.classification = options.classification;
    }
    if (options.visibility) {
      requestBody.visibility = options.visibility;
    }
    return requestBody;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const body: any = this.mapRequestBody(args.options);

    const requestOptions: any = {
      url: `${this.resource}/beta/groups/${encodeURIComponent(args.options.teamId)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      body: body,
      json: true
    };

    request
      .patch(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Microsoft Teams team for which to update settings'
      },
      {
        option: '--displayName [displayName]',
        description: 'The display name for the Microsoft Teams team'
      },
      {
        option: '--description [description]',
        description: 'The description for the Microsoft Teams team'
      },
      {
        option: '--mailNickName [mailNickName]',
        description: 'The mail alias for the Microsoft Teams team'
      },
      {
        option: '--classification [classification]',
        description: 'The classification for the Microsoft Teams team'
      },
      {
        option: '--visibility [visibility]',
        description: 'The visibility of the Microsoft Teams team. Valid values Private|Public',
        autocomplete: ['Private', 'Public']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (args.options.visibility) {
        if (args.options.visibility.toLowerCase() !== 'private' && args.options.visibility.toLowerCase() !== 'public') {
          return `${args.options.visibility} is not a valid visibility type. Allowed values are Private|Public`;
        }
      }

      return true;
    };
  }
}

module.exports = new TeamsSetCommand();