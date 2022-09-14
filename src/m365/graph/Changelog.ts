export interface Changelog {
  title: string;
  url: string;
  description: string;
  items: ChangelogItem[];
}

export interface ChangelogItem {
  guid: string;
  category: string;
  title: string;
  description: string;
  pubDate: Date;
}