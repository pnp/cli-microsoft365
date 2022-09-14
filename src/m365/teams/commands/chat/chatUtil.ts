import { AadUserConversationMember, Chat, ConversationMember } from '@microsoft/microsoft-graph-types';
import { odata } from '../../../../utils/odata';

export const chatUtil = {
  
  /**
   * Finds existing Microsoft Teams chats by participants, using the Microsoft Graph
   * @param expectedMemberEmails a string array of participant emailaddresses   
   * @param logger a logger to pipe into the graph request odata helper.
   */
  async findExistingChatsByParticipants(expectedMemberEmails: string[]): Promise<Chat[]> {
    const chatType = expectedMemberEmails.length === 2 ? 'oneOnOne' : 'group';
    const endpoint = `https://graph.microsoft.com/v1.0/chats?$filter=chatType eq '${chatType}'&$expand=members&$select=id,topic,createdDateTime,members`;
    const foundChats: Chat[] = [];
    
    const chats = await odata.getAllItems<Chat>(endpoint);

    for (const chat of chats) {
      const chatMembers = chat.members as ConversationMember[];
      if (chatMembers.length === expectedMemberEmails.length) {
        const chatMemberEmails = chatMembers.map(member => (member as AadUserConversationMember).email?.toLowerCase());

        if (expectedMemberEmails.every(email => chatMemberEmails.some(memberEmail => memberEmail === email))) {
          foundChats.push(chat);
        }
      }
    }

    return foundChats;
  },

  /**
   * Finds existing Microsoft Teams chats by name, using the Microsoft Graph
   * @param name the name of the chat conversation to find
   * @param logger a logger to pipe into the graph request odata helper.
   */
  async findExistingGroupChatsByName(name: string): Promise<Chat[]> {
    const endpoint = `https://graph.microsoft.com/v1.0/chats?$filter=topic eq '${encodeURIComponent(name).replace("'", "''")}'&$expand=members&$select=id,topic,createdDateTime,chatType`;
    return odata.getAllItems<Chat>(endpoint);    
  }
};