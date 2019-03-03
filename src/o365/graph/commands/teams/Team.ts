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
}