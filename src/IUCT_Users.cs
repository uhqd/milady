using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media;
using SW = System.Windows;

namespace catchDose
{
    public class IUCT_Users
    {
        private List<IUCT_User> _users_list;
        public IUCT_Users()
        {
            _users_list = new List<IUCT_User>();

            #region open and read xlsx file with users 


            string userListFilePath = Directory.GetCurrentDirectory() + @"\users\Users-IUCT.csv";

            if(!File.Exists(userListFilePath))
            {
                SW.MessageBox.Show("Le fichier des utilisateurs est introuvable :\n" + userListFilePath);
                return;
            }

            foreach (string line in File.ReadLines(userListFilePath))
            {
                // Séparer la ligne en colonnes (tu peux adapter le séparateur)
                string[] columns = line.Split(';'); // ou ',' selon ton fichier

                string id = columns[0];
                string firstname = columns[1];
                string lastname = columns[2];
                string sex = columns[3];
                string function = columns[4];
                string mybgcolor = columns[4];
                string myfgcolor = columns[4];

                if ((id != "") && (id != null))
                {
                    IUCT_User myUser = new IUCT_User() { userId = id, UserFirstName = firstname, UserFamilyName = lastname, Gender = sex, Function = function, UserBackgroundColor = mybgcolor, UserForeGroundColor = myfgcolor };
                    _users_list.Add(myUser);
                }
            }







        }



        #endregion






        public List<IUCT_User> UsersList
        {
            get { return _users_list; }
            set { _users_list = value; }
        }

    }
}
