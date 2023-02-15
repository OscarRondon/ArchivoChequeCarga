using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM.Framework;

namespace ArchivoChequeCarga.Menu
{
    class Menu
    {

        #region Events
        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                switch (pVal.BeforeAction)
                {
                    case true:
                        switch (pVal.MenuUID)
                        {
                            case "ArchivoChequeCarga.ArchivoCheques":
                                try
                                {
                                    Application.SBO_Application.Forms.GetForm("ArchivoChequeCarga.ArchivoCheques", 0).Select();
                                }
                                catch
                                {
                                    Forms.ArchivoCheques activeForm = new Forms.ArchivoCheques();
                                    activeForm.Show();
                                }
                                break;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: ArchivoChequeCarga.Menu.Menu.cs > SBO_Application_MenuEvent(): " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Methods
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

            try
            {
                oMenuItem = Application.SBO_Application.Menus.Item("43538"); // Banks > Outgoing payments'
                if (!oMenuItem.SubMenus.Exists("ArchivoChequeCarga.ArchivoCheques"))
                {
                    oMenus = oMenuItem.SubMenus;

                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "ArchivoChequeCarga.ArchivoCheques";
                    oCreationPackage.String = "Archivo de cheques para carga";
                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: ArchivoChequeCarga.Menu.Menu.cs > AddMenuItems(): " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion
    }
}
