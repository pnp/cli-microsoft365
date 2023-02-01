import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';


interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  rosterId: string;
}

class PlannerRosterMemberListCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_MEMBER_LIST;
  }

  public get description(): string {
    return 'Lists members of the specified Microsoft Planner Roster';
  }

  constructor() {
    super();

    this.#initOptions();
  }


  #initOptions(): void {
    this.options.unshift(
      {
        option: '--rosterId <rosterId>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr('Retrieving members of the specified Microsoft Planner Roster');
    }

    const url = `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members`;
    const response = await odata.getAllItems(url);
    logger.log(response);
  }

}


module.exports = new PlannerRosterMemberListCommand();
