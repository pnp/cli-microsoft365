import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const data: any = {
      funSettings: {}
    };
    TeamsFunSettingsSetCommand.booleanProps.forEach(p => {
      if (typeof (args.options as any)[p] !== 'undefined') {
        data.funSettings[p] = (args.options as any)[p] === 'true';
      }
    });

    if (args.options.giphyContentRating) {
      data.funSettings.giphyContentRating = args.options.giphyContentRating;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: data,
      responseType: 'json'
    };

    request
      .patch(requestOptions)
      .then((): void => {
        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>'
      },
      {
        option: '--allowGiphy [allowGiphy]'
      },
      {
        option: '--giphyContentRating [giphyContentRating]'
      },
      {
        option: '--allowStickersAndMemes [allowStickersAndMemes]'
      },
      {
        option: '--allowCustomMemes [allowCustomMemes]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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
  }
}

module.exports = new TeamsFunSettingsSetCommand();