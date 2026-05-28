import { DOMParser } from '@xmldom/xmldom';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { md } from '../../../../utils/md.js';
import { validation } from '../../../../utils/validation.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import { Changelog, ChangelogItem } from '../../Changelog.js';
import commands from '../../commands.js';

const allowedVersions = ['beta', 'v1.0'];
const allowedChangeTypes = ['Addition', 'Change', 'Deletion', 'Deprecation'];
const allowedServices = [
  'Applications', 'Calendar', 'Change notifications', 'Cloud communications',
  'Compliance', 'Cross-device experiences', 'Customer booking', 'Device and app management',
  'Education', 'Files', 'Financials', 'Groups',
  'Identity and access', 'Mail', 'Notes', 'Notifications',
  'People and workplace intelligence', 'Personal contacts', 'Reports', 'Search',
  'Security', 'Sites and lists', 'Tasks and plans', 'Teamwork',
  'To-do tasks', 'Users', 'Workbooks and charts'
];

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  versions: z.string().optional().alias('v'),
  changeType: z.string().optional().alias('c'),
  services: z.string().optional().alias('s'),
  startDate: z.string().optional(),
  endDate: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphChangelogListCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CHANGELOG_LIST;
  }

  public get description(): string {
    return 'Gets an overview of specific API-level changes in Microsoft Graph v1.0 and beta';
  }

  public defaultProperties(): string[] | undefined {
    return ['category', 'title', 'description'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => {
        if (!options.versions) {
          return true;
        }
        return !options.versions.toLocaleLowerCase().split(',').some(x => !allowedVersions.map(y => y.toLocaleLowerCase()).includes(x));
      }, {
        error: `The verions contains an invalid value. Specify either ${allowedVersions.join(', ')} as properties`,
        path: ['versions']
      })
      .refine(options => {
        if (!options.changeType) {
          return true;
        }
        return allowedChangeTypes.map(x => x.toLocaleLowerCase()).includes(options.changeType.toLocaleLowerCase());
      }, {
        error: `The change type contain an invalid value. Specify either ${allowedChangeTypes.join(', ')} as properties`,
        path: ['changeType']
      })
      .refine(options => {
        if (!options.services) {
          return true;
        }
        return !options.services.toLocaleLowerCase().split(',').some(x => !allowedServices.map(y => y.toLocaleLowerCase()).includes(x));
      }, {
        error: `The services contains invalid value. Specify either ${allowedServices.join(', ')} as properties`,
        path: ['services']
      })
      .refine(options => !options.startDate || validation.isValidISODate(options.startDate), {
        error: 'The startDate is not a valid ISO date string',
        path: ['startDate']
      })
      .refine(options => !options.endDate || validation.isValidISODate(options.endDate), {
        error: 'The endDate is not a valid ISO date string',
        path: ['endDate']
      })
      .refine(options => !(options.endDate && options.startDate && new Date(options.endDate) < new Date(options.startDate)), {
        error: 'The endDate should be later than startDate',
        path: ['endDate']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const allowedChangeType = args.options.changeType && allowedChangeTypes.find(x => x.toLocaleLowerCase() === args.options.changeType!.toLocaleLowerCase());
      const searchParam = args.options.changeType ? `/?filterBy=${allowedChangeType}` : '';

      const requestOptions: CliRequestOptions = {
        url: `https://developer.microsoft.com/en-us/graph/changelog/rss${searchParam}`,
        headers: {
          'accept': 'text/xml',
          'x-anonymous': 'true'
        }
      };

      const output: any = await request.get(requestOptions);
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(output.toString(), "text/xml");

      const changelog = this.filterThroughOptions(args.options, this.mapChangelog(xmlDoc, args));

      await logger.log(changelog.items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private filterThroughOptions(options: Options, changelog: Changelog): Changelog {
    let items: ChangelogItem[] = changelog.items;

    if (options.services) {
      const matchedServices: string[] = allowedServices
        .filter(allowedService => options.services!.toLocaleLowerCase().split(',').includes(allowedService.toLocaleLowerCase()));

      items = changelog.items.filter(item => matchedServices.includes(item.title));
    }

    if (options.versions) {
      const matchedVersions: string[] = allowedVersions
        .filter(allowedVersion => options.versions!.toLocaleLowerCase().split(',').includes(allowedVersion.toLocaleLowerCase()));

      items = items.filter(item => matchedVersions.includes(item.category));
    }

    if (options.startDate) {
      const startDate: Date = new Date(options.startDate);

      items = items.filter(item => item.pubDate >= startDate);
    }

    if (options.endDate) {
      const endDate: Date = new Date(options.endDate);

      items = items.filter(item => item.pubDate <= endDate);
    }

    // Make sure everything is unique based on the item guid
    items = [...new Map(items.map((item) => [item.guid, item])).values()];
    changelog.items = items.sort((itemA, itemB) => Number(itemB.pubDate) - Number(itemA.pubDate));

    return changelog;
  }

  private mapChangelog(xmlDoc: any, args: CommandArgs): Changelog {
    const channel = xmlDoc.getElementsByTagName('channel').item(0);

    const changelog: Changelog = {
      title: channel.getElementsByTagName('title').item(0).textContent,
      description: channel.getElementsByTagName('description').item(0).textContent,
      url: channel.getElementsByTagName('link').item(0).textContent,
      items: []
    } as Changelog;

    Array.from(xmlDoc.getElementsByTagName('item')).forEach((item: any) => {
      const description: string = cli.shouldTrimOutput(args.options.output) ?
        md.md2plain(item.getElementsByTagName('description').item(0).textContent, '') :
        item.getElementsByTagName('description').item(0).textContent;

      changelog.items.push({
        guid: item.getElementsByTagName('guid').item(0).textContent,
        category: item.getElementsByTagName('category').item(1).textContent,
        title: item.getElementsByTagName('title').item(0).textContent,
        description: cli.shouldTrimOutput(args.options.output) ?
          description.length > 50 ? `${description.substring(0, 47)}...` : description :
          description,
        pubDate: new Date(item.getElementsByTagName('pubDate').item(0).textContent)
      } as ChangelogItem);
    });

    return changelog;
  }
}

export default new GraphChangelogListCommand();