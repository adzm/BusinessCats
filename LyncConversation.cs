using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessCats
{
    public class LyncConversation
    {
        public void UpdateDesc()
        {
            string people = "";
            foreach (var participant in conversation.Participants)
            {
                if (participant.IsSelf)
                {
                    continue;
                }

                var name = participant.Contact.GetContactInformation(ContactInformationType.DisplayName);
                people += name.ToString() + "; ";
            }

            people = people.TrimEnd(';', ' ');
            desc = people;
        }

        public string desc;
        public Conversation conversation;
        public Dictionary<Participant, LyncParticipant> participants;

        public LyncParticipant FindParticipant(Participant participant)
        {
            if (participants.ContainsKey(participant))
            {
                return participants[participant];
            }

            return null;
        }

        public override string ToString()
        {
            return desc;
        }
    }
}
