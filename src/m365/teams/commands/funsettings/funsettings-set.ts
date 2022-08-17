import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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
    return commands.FUNSETTINGS_SET;
  }

  public get description(): string {
    return 'Updates fun settings of a Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        giphyContentRating: args.options.giphyContentRating
      });
      TeamsFunSettingsSetCommand.booleanProps.forEach(p => {
        this.telemetryProperties[p] = (args.options as any)[p];
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.teamId)) {
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
            return `giphyContentRating value ${value} is not valid.  Please specify Strict or Moderate.`;
          }
        }

        return true;
      }
    );
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
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new TeamsFunSettingsSetCommand();