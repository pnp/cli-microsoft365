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
  id?: string;
  teamId?: string;
  name?: string;
  displayName?: string;
  description?: string;
  mailNickName?: string;
  classification?: string;
  visibility?: string;
}

class TeamsTeamSetCommand extends GraphCommand {
  private static props: string[] = [
    'displayName',
    'description',
    'mailNickName',
    'classification',
    'visibility '
  ];

  public get name(): string {
    return commands.TEAM_SET;
  }

  public get description(): string {
    return 'Updates settings of a Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      TeamsTeamSetCommand.props.forEach((p: string) => {
        this.telemetryProperties[p] = typeof (args.options as any)[p] !== 'undefined';
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '--teamId [teamId]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--displayName [displayName]'
      },
      {
        option: '--description [description]'
      },
      {
        option: '--mailNickName [mailNickName]'
      },
      {
        option: '--classification [classification]'
      },
      {
        option: '--visibility [visibility]',
        autocomplete: ['Private', 'Public']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
	      return `${args.options.teamId} is not a valid GUID`;
	    }

	    if (args.options.id && !validation.isValidGuid(args.options.id)) {
	      return `${args.options.id} is not a valid GUID`;
	    }

	    if (args.options.visibility) {
	      if (args.options.visibility.toLowerCase() !== 'private' && args.options.visibility.toLowerCase() !== 'public') {
	        return `${args.options.visibility} is not a valid visibility type. Allowed values are Private|Public`;
	      }
	    }

	    return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'teamId']);
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};
    if (options.name) {
      requestBody.displayName = options.name;
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.teamId) {
      args.options.id = args.options.teamId;

      this.warn(logger, `Option 'teamId' is deprecated. Please use 'id' instead.`);
    }

    if (args.options.displayName) {
      args.options.name = args.options.displayName;

      this.warn(logger, `Option 'displayName' is deprecated. Please use 'name' instead.`);
    }

    const data: any = this.mapRequestBody(args.options);

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups/${encodeURIComponent(args.options.id as string)}`,
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

module.exports = new TeamsTeamSetCommand();