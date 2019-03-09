export interface Channel {
  id: string;
  displayName: string | null;
  description: string | null;
  isFavoriteByDefault:boolean|null;
  email:string|null;
  webUrl:string|null;
}