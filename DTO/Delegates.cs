using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace emedit.DTO
{
    public delegate void UserSendEvent(UserDTO dto);
    internal class Delegates
    {
        // 로그인한 사용자의 정보를 전달하는 class
        public static UserSendEvent UserSendEvent;

        public static UserDTO userdto = new UserDTO();

        public void SetUser(UserDTO dto)
        {
            userdto = dto;
            //MessageBox.Show("dto 전달완료" + userdto.usrid);
            //Console.WriteLine("dto 전달완료 " + userdto.usrid);
        }

        public static void usersend2()
        {
            Form2 f2 = new Form2();
            UserSendEvent += new UserSendEvent(f2.SetUser);
            UserSendEvent(userdto);
        }
        public static void usersend4()
        {
            Form4 f2 = new Form4();
            UserSendEvent += new UserSendEvent(f2.SetUser);
            UserSendEvent(userdto);
        }
        public static void usersend6()
        {
            Form6 f2 = new Form6();
            UserSendEvent += new UserSendEvent(f2.SetUser);
            UserSendEvent(userdto);
        }

        public static void usersend10()
        {
            Form10 f2 = new Form10();
            UserSendEvent += new UserSendEvent(f2.SetUser);
            UserSendEvent(userdto);
        }


    }
}
