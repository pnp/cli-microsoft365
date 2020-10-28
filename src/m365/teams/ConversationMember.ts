export interface ConversationMember {
  id: string;
  roles: string[];
  displayName: string | null;
  userId: string | null;
  email: string | null;
}