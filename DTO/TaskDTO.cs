using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace emedit.DTO
{
    public class TaskDTO
    {
        
        // task_inf 테이블의 데이터를 저장하는 class
        
        // teskid
        public string taskid {  get; set; }
        // 프로젝트 no
        public string prjtno { get; set; }
        // 프로젝트 id
       // public string prjid { get; set; }
        // 접수일자
        public string regdt { get; set; }
        // 작성자
        public string userid { get; set; }
        // 상태
        public string state { get; set; }
        // 공개등급
        public string level { get; set; }
        // 원본태스크 id
        public string orgtaskid { get; set; }
        // 작성자
        public string wrtuserid { get; set; }
        // 작성일자
        public string wrtdt { get; set; }
        // 제목
        public string head { get; set; }
        // 내용
        public string reqinfo { get; set; }
        // 요청구분
        public string receiptflg { get; set; }
        // 요청자
        public string reqnm { get; set; }

        // 이메일
        public string email { get; set; }   

        // 담당자이메일
        public string usremail { get; set; }
    }
}
