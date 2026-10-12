import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  scope: z.enum(['Site', 'Web', 'All']).optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCustomActionListCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_LIST;
  }

  public get description(): string {
    return 'Lists all user custom actions at the given scope';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['Name', 'Location', 'Scope', 'Id'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const scope: string = args.options.scope ? args.options.scope : 'All';

      if (this.debug) {
        await logger.logToStderr(`Attempt to get custom actions list with scope: ${scope}`);
        await logger.logToStderr('');
      }

      const customActions = await spo.getCustomActions(args.options.webUrl, args.options.scope);

      if (args.options.output !== 'json') {
        customActions.forEach(a => a.Scope = this.humanizeScope(a.Scope) as any);
      }

      await logger.log(customActions);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private humanizeScope(scope: number): string {
    switch (scope) {
      case 2:
        return "Site";
      case 3:
        return "Web";
    }

    return `${scope}`;
  }
}

export default new SpoCustomActionListCommand();