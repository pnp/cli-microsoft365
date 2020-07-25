import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import request from '../../../../request';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  allowGiphy: string;
  giphyContentRating: string;
  allowStickersAndMemes: string;
  allowCustomMemes: string;
}

class TeamsFunSettingsSetCommand extends GraphCommand {
  private static booleanProps: string[] = [
    'allowGiphy',
    'allowStickersAndMemes',
    'allowCustomMemes'
  ];

  public get name(): string {
    return `${commands.TEAMS_FUNSETTINGS_SET}`;
  }

  public get description(): string {
    return 'Updates fun settings of a Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    TeamsFunSettingsSetCommand.booleanProps.forEach(p => {
      telemetryProps[p] = (args.options as any)[p];
    });
    telemetryProps.giphyContentRating = args.options.giphyContentRating;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const body: any = {
      funSettings: {}
    };
    TeamsFunSettingsSetCommand.booleanProps.forEach(p => {
      if (typeof (args.options as any)[p] !== 'undefined') {
        body.funSettings[p] = (args.options as any)[p] === 'true';
      }
    });

    if (args.options.giphyContentRating) {
      body.funSettings.giphyContentRating = args.options.giphyContentRating;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}`,
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
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Teams team for which to update settings'
      },
      {
        option: '--allowGiphy [allowGiphy]',
        description: 'Set to true to allow giphy and to false to disable it'
      },
      {
        option: '--giphyContentRating [giphyContentRating]',
        description: 'Settings to set content rating for giphy. Allowed values Strict|Moderate'
      },
      {
        option: '--allowStickersAndMemes [allowStickersAndMemes]',
        description: 'Set to true to allow stickers and memes and to false to disable them'
      },
      {
        option: '--allowCustomMemes [allowCustomMemes]',
        description: 'Set to true to allow custom memes and to false to disable them'
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

      let isValid: boolean = true;
      let value, property: string = '';
      TeamsFunSettingsSetCommand.booleanProps.every(p => {
        property = p;
        value = (args.options as any)[p];
        isValid = typeof value === 'undefined' ||
          value === 'true' ||
          value === 'false';
        return isValid;
      });

      if (!isValid) {
        return `Value ${value} for option ${property} is not a valid boolean`;
      }

      if (args.options.giphyContentRating) {
        const giphyContentRating = args.options.giphyContentRating.toLowerCase();
        if (giphyContentRating !== 'strict' && giphyContentRating !== 'moderate') {
          return `giphyContentRating value ${value} is not valid.  Please specify Strict or Moderate.`
        }
      }

      return true;
    };
  }
}

module.exports = new TeamsFunSettingsSetCommand();