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
  partsToClone: string;
  description?: string;
  classification?: string;
  visibility?: string;
}

class TeamsTeamCloneCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_CLONE;
  }

  public get description(): string {
    return 'Creates a clone of a Microsoft Teams team';
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
      Object.assign(this.telemetryProperties, {
        description: typeof args.options.description !== 'undefined',
        classification: typeof args.options.classification !== 'undefined',
        visibility: typeof args.options.visibility !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        teamId: typeof args.options.teamId !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined'
        
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [teamId]'
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
        option: '-p, --partsToClone <partsToClone>',
        autocomplete: ['apps', 'channels', 'members', 'settings', 'tabs']
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-c, --classification [classification]'
      },
      {
        option: '-v, --visibility [visibility]',
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

	    const partsToClone: string[] = args.options.partsToClone.replace(/\s/g, '').split(',');
	    for (const partToClone of partsToClone) {
	      const part: string = partToClone.toLowerCase();
	      if (part !== 'apps' &&
	        part !== 'channels' &&
	        part !== 'members' &&
	        part !== 'settings' &&
	        part !== 'tabs') {
	        return `${part} is not a valid partsToClone. Allowed values are apps|channels|members|settings|tabs`;
	      }
	    }

	    if (args.options.visibility) {
	      const visibility: string = args.options.visibility.toLowerCase();

	      if (visibility !== 'private' &&
	        visibility !== 'public') {
	        return `${args.options.visibility} is not a valid visibility type. Allowed values are Private|Public`;
	      }
	    }

	    return true;
      }
    );
  }

  #initOptionSets(): void {
  	this.optionSets.push(
  	  ['id', 'teamId'],
      ['name', 'displayName']
  	);
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

    const data: any = {
      displayName: args.options.name,
      mailNickname: this.generateMailNickname(args.options.name as string),
      partsToClone: args.options.partsToClone
    };

    if (args.options.description) {
      data.description = args.options.description;
    }

    if (args.options.classification) {
      data.classification = args.options.classification;
    }

    if (args.options.visibility) {
      data.visibility = args.options.visibility;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.id as string)}/clone`,
      headers: {
        "content-type": "application/json",
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: data
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  /**
   * There is a know issue that the mailNickname is currently ignored and cannot be set by the user
   * However the mailNickname is still required by the payload so to deliver better user experience
   * the CLI generates mailNickname for the user 
   * so the user does not have to specify something that will be ignored.
   * For more see: https://docs.microsoft.com/en-us/graph/api/team-clone?view=graph-rest-1.0#request-data
   * This method has to be removed once the graph team fixes the issue and then the actual value
   * of the mailNickname would have to be specified by the CLI user.
   * @param displayName teams display name
   */
  private generateMailNickname(displayName: string): string {
    // currently the Microsoft Graph generates mailNickname in a similar fashion
    return `${displayName.replace(/[^a-zA-Z0-9]/g, "")}${Math.floor(Math.random() * 9999)}`;
  }
}

module.exports = new TeamsTeamCloneCommand();