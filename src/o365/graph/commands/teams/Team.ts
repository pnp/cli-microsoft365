export interface Team {
  id: string;
  displayName: string;
  description: string;
  isArchived: boolean | undefined;
  messagingSettings?: {
    allowUserEditMessages: boolean,
    allowUserDeleteMessages: boolean,
    allowOwnerDeleteMessages: boolean,
    allowTeamMentions: boolean,
    allowChannelMentions: boolean
  }
  memberSettings?: {
    allowCreateUpdateChannels: boolean,
    allowDeleteChannels: boolean,
    allowAddRemoveApps: boolean,
    allowCreateUpdateRemoveTabs: boolean,
    allowCreateUpdateRemoveConnectors: boolean
  }
}