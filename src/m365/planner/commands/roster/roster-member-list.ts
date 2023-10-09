import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
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
    this.#initTypes();
  }


  #initOptions(): void {
    this.options.unshift(
      {
        option: '--rosterId <rosterId>'
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('rosterId');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Retrieving members of the specified Microsoft Planner Roster');
    }

    try {
      const response = await odata.getAllItems(`${this.resource}/beta/planner/rosters/${args.options.rosterId}/members`);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerRosterMemberListCommand();
