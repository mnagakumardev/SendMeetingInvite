using System;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace ConsoleApp6
{
    class Program
    {
        static void Main(string[] args)
        {
            string startTime1 = Convert.ToDateTime(DateTime.Now.AddHours(1)).ToString("yyyyMMddTHHmmssZ");
            string endTime1 = Convert.ToDateTime(DateTime.Now.AddHours(2)).ToString("yyyyMMddTHHmmssZ");
            var sc = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,

                Credentials = new NetworkCredential("fromemailaddress@gmail.com", "Password")
            };


            MailMessage msg = new MailMessage();

            msg.From = new MailAddress("FromEmailAddress@gmail.com", "My Self Service");
            msg.To.Add(new MailAddress("ToEmailAddress@outlook.com"));
            msg.Subject = "Holiday Approval";

            string mailBody = "";
            mailBody += "<table style='width: 600px; height: 900px; background-color: #E2BB7E;'>";
            mailBody += "<tr><td style='width: 581px; height: 147px;'>";
            mailBody += "<img src='cid:TopBanner' />";
            mailBody += "</td></tr>";
            mailBody += "<tr><td style='background-color: #fff; width: 581px; height: 234px;'>";
            mailBody += "<table><tr><td style='width: 290px; padding-left: 10px;'>";
            mailBody += "<img src='cid:dealer_logo' />";
            mailBody += "<br />";
            mailBody += "<br />";
            mailBody += "Dealer Address";
            mailBody += "<br />";
            mailBody += "City, State ZIP";
            mailBody += "<br />";
            mailBody += "XXX-XXX-XXXX";
            mailBody += "<br />";
            mailBody += "<a href='http://www.abc.com' target='_blank' style='background-color: #f44336;color: white;padding: 14px 25px;text-align: center;text-decoration: none;display: inline-block;' > www.abc.com </a> ";
            mailBody += "</td><td style='font-size: 44px; font-weight: bold; width: 291px; border: dashed 1px #000000; padding: 5px 5px 5px 5px; text-align: center;'>";
            mailBody += "$300";
            mailBody += "<br />";
            mailBody += "Rebate";
            mailBody += "<br />";
            mailBody += "<a href='#' target='_blank' style='font-size: 20px;'>Click here</a>";
            mailBody += "<span style='font-size: 20px;'>";
            mailBody += " for coupon";
            mailBody += "</span>";
            mailBody += "</td></tr></table></td></tr>";
            mailBody += "<tr style='background-color: #47AA42; color: #fff; font-size: 22px; width: 581px;'><td style='padding-left: 5px;'>";
            mailBody += "It's your last chance to save even more with a $1,500 tax credit.*";
            mailBody += "<br />";
            mailBody += "<span style='font-size: 15px; padding-left: 200px;'>";
            mailBody += "</span>";
            mailBody += "</td></tr>";
            mailBody += "<tr><td style='width: 581px; height: 455px;'>";
            mailBody += "<img src='cid:BottomImage' />";
            mailBody += "</td></tr>";
            mailBody += "<tr><td style='width: 590px; height: 80px; font-size: 0.6em; padding-left: 10px; font-weight: bold;'>";
            mailBody += "</td></tr>";
            mailBody += "</table>";

            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(mailBody.ToString(), new System.Net.Mime.ContentType("text/html"));


            msg.AlternateViews.Add(htmlView);

            msg.IsBodyHtml = true;


            StringBuilder str = new StringBuilder();
            str.AppendLine("BEGIN:VCALENDAR");

            str.AppendLine("PRODID:-//A//Outlook MIMEDIR//EN");
            str.AppendLine("VERSION:2.0");
            str.AppendLine("METHOD:REQUEST");

            str.AppendLine("BEGIN:VEVENT");

            str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", startTime1));
            str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
            str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", endTime1));
            str.AppendLine(string.Format("LOCATION: {0}", "Location"));

            // UID should be unique.
            str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
            str.AppendLine(string.Format("DESCRIPTION:{0}", msg.Body));
            str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", msg.Body));
            str.AppendLine(string.Format("SUMMARY:{0}", msg.Subject));

            str.AppendLine("STATUS:CONFIRMED");
            str.AppendLine("BEGIN:VALARM");
            str.AppendLine("TRIGGER:-PT15M");
            str.AppendLine("ACTION:Accept");
            str.AppendLine("DESCRIPTION:Reminder");
            str.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:BUSY");
            str.AppendLine("END:VALARM");
            str.AppendLine("END:VEVENT");

            str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", msg.From.Address));
            str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", msg.To[0].DisplayName, msg.To[0].Address));

            str.AppendLine("END:VCALENDAR");


            System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType("text/calendar");
            ct.Parameters.Add("method", "REQUEST");
            ct.Parameters.Add("name", "meeting.ics");
            AlternateView avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), ct);
            msg.AlternateViews.Add(avCal);

            sc.ServicePoint.MaxIdleTime = 2;

            sc.Send(msg);

        }
    }
}
