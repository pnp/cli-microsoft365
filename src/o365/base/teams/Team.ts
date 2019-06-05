export interface Team {
  id: string;
  displayName: string;
  description: string;
  isArchived: boolean | undefined;
  messagingSettings?: any;
  memberSettings?: any;
  guestSettings?: any;
}