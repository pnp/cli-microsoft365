import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { md, validation } from '../../../../utils';
import AnonymousCommand from '../../../base/AnonymousCommand';
import { Changelog, ChangelogItem } from '../../Changelog';
import commands from '../../commands';
import request from '../../../../request';
import { DOMParser } from '@xmldom/xmldom';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  versions?: string;
  changeType?: string;
  services?: string;
  startDate?: string;
  endDate?: string;
}

class GraphChangelogListCommand extends AnonymousCommand {
  private allowedVersions: string[] = ['beta', 'v1.0'];
  private allowedChangeTypes: string[] = ['Addition', 'Change', 'Deletion', 'Deprecation'];
  private allowedServices: string[] = [
    'Applications', 'Calendar', 'Change notifications', 'Cloud communications', 
    'Compliance', 'Cross-device experiences', 'Customer booking', 'Device and app management', 
    'Education', 'Files', 'Financials', 'Groups', 
    'Identity and access', 'Mail', 'Notes', 'Notifications', 
    'People and workplace intelligence', 'Personal contacts', 'Reports', 'Search', 
    'Security', 'Sites and lists', 'Tasks and plans', 'Teamwork', 
    'To-do tasks', 'Users', 'Workbooks and charts'
  ];

  public get name(): string {
    return commands.CHANGELOG_LIST;
  }

  public get description(): string {
    return 'Gets an overview of specific API-level changes in Microsoft Graph v1.0 and beta';
  }

  public defaultProperties(): string[] | undefined {
    return ['category', 'title', 'description'];
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
        versions: typeof args.options.versions !== 'undefined',
        changeType: typeof args.options.changeType !== 'undefined',
        services: typeof args.options.services !== 'undefined',
        startDate: typeof args.options.startDate !== 'undefined',
        endDate: typeof args.options.endDate !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-v, --versions [versions]', autocomplete: this.allowedVersions },
      { option: "-c, --changeType [changeType]", autocomplete: this.allowedChangeTypes },
      { option: "-s, --services [services]", autocomplete: this.allowedServices },
      { option: "--startDate [startDate]" },
      { option: "--endDate [endDate]" }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (
          args.options.versions && 
          args.options.versions.toLocaleLowerCase().split(',').some(x => !this.allowedVersions.map(y => y.toLocaleLowerCase()).includes(x))) {
          return `The verions contains an invalid value. Specify either ${this.allowedVersions.join(', ')} as properties`;
        }
    
        if (
          args.options.changeType && 
          !this.allowedChangeTypes.map(x => x.toLocaleLowerCase()).includes(args.options.changeType.toLocaleLowerCase())) {
          return `The change type contain an invalid value. Specify either ${this.allowedChangeTypes.join(', ')} as properties`;
        }
    
        if (
          args.options.services && 
          args.options.services.toLocaleLowerCase().split(',').some(x => !this.allowedServices.map(y => y.toLocaleLowerCase()).includes(x))) {
          return `The services contains invalid value. Specify either ${this.allowedServices.join(', ')} as properties`;
        }
    
        if (args.options.startDate && !validation.isValidISODate(args.options.startDate)) {
          return 'The startDate is not a valid ISO date string';
        }
    
        if (args.options.endDate && !validation.isValidISODate(args.options.endDate)) {
          return 'The endDate is not a valid ISO date string';
        }
    
        if (args.options.endDate && args.options.startDate && new Date(args.options.endDate) < new Date(args.options.startDate)) {
          return 'The endDate should be later than startDate';
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const allowedChangeType = args.options.changeType && this.allowedChangeTypes.find(x => x.toLocaleLowerCase() === args.options.changeType!.toLocaleLowerCase());
    const searchParam = args.options.changeType ? `/?filterBy=${allowedChangeType}` : '';

    const requestOptions: any = {
      url: `https://developer.microsoft.com/en-us/graph/changelog/rss${searchParam}`,
      headers: {
        'accept': 'text/xml',
        'x-anonymous': 'true'
      }
    };

    request
      .get(requestOptions)
      .then((output: any) => {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(output.toString(), "text/xml");

        const changelog = this.filterThroughOptions(args.options, this.mapChangelog(xmlDoc, args));
        
        logger.log(changelog.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private filterThroughOptions(options: Options, changelog: Changelog): Changelog {
    let items: ChangelogItem[] = changelog.items;

    if (options.services) {
      const allowedServices: string[] = this.allowedServices
        .filter(allowedService => options.services!.toLocaleLowerCase().split(',').includes(allowedService.toLocaleLowerCase()));

      items = changelog.items.filter(item => allowedServices.includes(item.title));
    }

    if (options.versions) {
      const allowedVersions: string[] = this.allowedVersions
        .filter(allowedVersion => options.versions!.toLocaleLowerCase().split(',').includes(allowedVersion.toLocaleLowerCase()));

      items = items.filter(item => allowedVersions.includes(item.category));
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
      const description: string = args.options.output === 'text' ? 
        md.md2plain(item.getElementsByTagName('description').item(0).textContent, '') :
        item.getElementsByTagName('description').item(0).textContent;

      changelog.items.push({
        guid: item.getElementsByTagName('guid').item(0).textContent,
        category: item.getElementsByTagName('category').item(1).textContent,
        title: item.getElementsByTagName('title').item(0).textContent,
        description: args.options.output === 'text' ? 
          description.length > 50 ? `${description.substring(0, 47)}...` : description : 
          description,
        pubDate: new Date(item.getElementsByTagName('pubDate').item(0).textContent)
      } as ChangelogItem);
    });

    return changelog;
  }
}

module.exports = new GraphChangelogListCommand();