using emedit.DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace emedit.DAO
{
    public class sendmail
    {
        // 이메일 발송 시 실행되는 class 파일

        // 요청사항 접수 시 전송되는 Text
        string text = "님 , 요청사항 접수가 완료되었습니다.";
        public sendmail() {
           
        }
        // 요청사항 등록 시 전송되는 이메일
        public void sendtomail(TaskDTO dto)
        {
            MailMessage mail = new MailMessage();
            // 보내는사람
            mail.To.Add(dto.email);
            // 회신받을 주소
            mail.From = new MailAddress(dto.email);
            // 메일 내용
            mail.Body = dto.reqnm + text;
            // 제목
            mail.Subject = "[요청사항 접수 완료]";
            mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
            // 한글 전송을 위해 메일 제목을 UTF-8로 인코딩 
            mail.SubjectEncoding = Encoding.UTF8;
            // 한글 전송을 위해 메일 내용을 UTF-8로 인코딩
            mail.BodyEncoding = Encoding.UTF8;
           
            // smtp로 전송하기위해 smtp client 객체를 생성
            SmtpClient smtp = smtpclient();

            try
            {
                // smtp 객체를 통해 mail 발송
                smtp.Send(mail);
                // mail 객체에 사용한 리소스를 모두 해지
                mail.Dispose();
                // 전송이 완료되면 MessageBox에 전송완료 알림창을 나타나게 설정
                MessageBox.Show("전송완료", "전송 완료");
            }
            catch (Exception ex)
            {
                // 메일 발송 실패시 오류 메시지 출력
                MessageBox.Show(ex.ToString());
                
            }
        }

        // stmp 설정 함수
        private SmtpClient smtpclient()
        {
            // SmtpClient 사용을 위한 smtp 객체 생성
            SmtpClient smtp = new SmtpClient();
            // smtp 메일 서버 주소 입력
            smtp.Host = "smtp.gmail.com";
            // smtp 메일 포트 주소 입력
            smtp.Port = 587;
           // stmp 메일 timeout 설정
            smtp.Timeout = 10000;
            // 서버 기본 인증 이용
            smtp.UseDefaultCredentials = true;
            // smtp SSL 보안 설정
            smtp.EnableSsl = true;
            // 이메일을 네트워크를 통해 SMTP 서버로 전송
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            // 사용자 아이디와 비밀번호
            smtp.Credentials = new System.Net.NetworkCredential("jm0729a", "nfswjacwoiscsylz");
            // 설정한 smtp 객체 반환
            return smtp;
        }
        
        // 작업결과 회신 시 전송되는 이메일
        public void sendtorev(string email, string text, string filenames)
        {
            MailMessage mail = new MailMessage();
            // 보내는사람
            mail.To.Add(email);
            // 회신받을 주소
            mail.From = new MailAddress(email);
            // 메일 내용
            mail.Body = text;
            // 제목
            mail.Subject = "[요청사항 처리 완료]";
            // 메일 발송에 실패하면 알림을 띄우게 설정
            mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
            // 한글 사용을 위해 메일 제목 인코딩은 UTF-8
            mail.SubjectEncoding = Encoding.UTF8;
            // 한글 사용을 위해 메일 내용 인코딩을 UTF-8로 설정
            mail.BodyEncoding = Encoding.UTF8;

            if(filenames != "x")

            {   // theexcelfilename : 첨부파일의 전체 경로
                Attachment theexcel = new Attachment(filenames);
            // theexcelfile_short_name : 메일에서 보이게 되는 첨부파일명

            mail.Attachments.Add(theexcel);

            }
            //MessageBox.Show(theexcel.Name);
            SmtpClient smtp = smtpclient();

            try
            {
                // smtp 객체를 통해 mail 발송
                smtp.Send(mail);
                // mail에 저장된 리소스를 모두 해지
                mail.Dispose();

                MessageBox.Show("전송완료", "전송 완료");
            }
            catch (Exception ex)
            {
                // 메일 발송 실패시 오류 메시지 출력
                MessageBox.Show(ex.ToString());
               
            }
        }
    }

}

