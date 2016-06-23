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
            desc = Describe(conversation);
        }
        
        public string Describe()
        {
            return Describe(conversation);
        }

        public static string Describe(Conversation conversation)
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

            string subject = "";
            {
                object subjectObj;
                if (conversation.Properties.TryGetValue(ConversationProperty.Subject, out subjectObj))
                {
                    if (subjectObj is string)
                    {
                        subject = (string)subjectObj;
                    }
                    else
                    {
                        subject = subjectObj.ToString();
                    }
                }
            }

            if (!string.IsNullOrEmpty(subject))
            {
                return $"{subject} ({people})";
            }
            else
            {
                return people;
            }
        }

        public string DescribeAsTitle()
        {
            return DescribeAsTitle(conversation);
        }

        public static string DescribeAsTitle(Conversation conversation)
        {
            int participantCount = conversation.Participants.Count;

            if (participantCount == 2)
            {
                foreach (var p in conversation.Participants)
                {
                    if (p.IsSelf)
                    {
                        continue;
                    }

                    object val;
                    if (p.Properties.TryGetValue(ParticipantProperty.Name, out val) && val is string)
                    {
                        return val as string;
                    }
                }
            }

            string participants = "";
            {
                string suffix = (participantCount == 1) ? "" : "s";
                participants = $"({participantCount} Participant{suffix})";
            }
            
            string subject = "";
            {
                object subjectObj;
                if (conversation.Properties.TryGetValue(ConversationProperty.Subject, out subjectObj))
                {
                    if (subjectObj is string)
                    {
                        subject = (string)subjectObj;
                    }
                    else
                    {
                        subject = subjectObj.ToString();
                    }
                }
            }

            if (string.IsNullOrEmpty(subject))
            {
                subject = "Conversation";
            }

            return $"{subject} {participants}";
        }

        // Subject can't be updated in new Meet Now conversation
        // Subject can't be updated once conversation in progress
        // Seems to only work for single-participant conversations, which then converts 
        // to normal conversation with the subject when more participants added... at 
        // which point you can't change the subject any more.
        // ConversationManager.AddConversation is different than Meet Now
        // a Conversation with only yourself will not work right
        // a Conversation only shows up when a message is sent
        public async static Task DoUpdateSubject(Conversation conversation, string subject)
        {
            try
            {
                await Task.Factory.FromAsync(conversation.BeginSetProperty, conversation.EndSetProperty, ConversationProperty.Subject, subject, conversation);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.ToString());
            }
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
