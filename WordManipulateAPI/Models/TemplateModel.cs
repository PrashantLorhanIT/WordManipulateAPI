using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WordManipulateAPI.Models
{
    public class LetterTemplateModel
    {
        public string RelativeTemplatePath { get; set; }
        public string Reference { get; set; }
        public DateTime CorrespondentDate { get; set; }      
        public string Address { get; set; }
        public string RecipientName { get; set; }       
        public string Subject { get; set; }
        public string Body { get; set; }
        public string SenderName { get; set; }
        public string SenderDesignation { get; set; }

        public string Enclosure { get; set; }
      
    }

    public class CircularTemplateModel
    {
        public string RelativeTemplatePath { get; set; }
        public string Reference { get; set; }
        public DateTime CorrespondentDate { get; set; }      
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Mdl { get; set; }
        public string AdHoc { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }
        public string Approver { get; set; }       
    }

    public class MemoTemplateModel
    {
        public string RelativeTemplatePath { get; set; }
        public string Reference { get; set; }
        public DateTime CorrespondentDate { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Mdl { get; set; }
        public string AdHoc { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }
        public string Approver { get; set; }
    }

    public enum TemplateType
    {
        Letter,
        Circular,
        Memo
    }

    public class MomTemplateModel
    {

        public string DocumentNumber { get; set; }
        public DateTime DueDate { get; set; }
        // public DateTime Time { get; set; }
        public string Location { get; set; }
        public string Subject { get; set; }
        public string MinutesRecorder { get; set; }
        // public string Attendees { get; set; }

        public string Organizer { get; set; }
        public string MeetingAgenda { get; set; }

        public List<AttendeesTemplateModel> attendee = new List<AttendeesTemplateModel>();

        public List<TaskCommentTemplateModel> TaskCmtaction = new List<TaskCommentTemplateModel>();

    }

    public class AttendeesTemplateModel
    {
        //ridMom: number; // Added
        public string name { get; set; }
        public string title { get; set; }
        public string unitorganization { get; set; }

        public string isPresent { get; set; }
    }

    public class RFITemplateModel
    {
        public string DocumentNumber { get; set; }
        public DateTime DueDate { get; set; }
        public DateTime RfiDate { get; set; }
        public string recipient { get; set; }
        public string Subject { get; set; }
        public string rfiRequestQuery { get; set; }
        public string Originator { get; set; }
        public string ridContract { get; set; }
        public string ridEntitySender { get; set; }
        public string ridEntityRecipient { get; set; }
        public string isConfidential { get; set; }

        public string response { get; set; }

    }

    public class TaskCommentTemplateModel
    {
        public string comments { get; set; }
        public string actionby { get; set; }
        //public List<actionby> actionby = new List<actionby>();
        public string subject { get; set; }
        public string isMinute { get; set; }
        public string dueDate { get; set; }
    }

}