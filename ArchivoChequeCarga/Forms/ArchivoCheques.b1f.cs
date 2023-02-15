using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Xml.Linq;
using System.IO;

namespace ArchivoChequeCarga.Forms
{
    [FormAttribute("ArchivoChequeCarga.Forms.ArchivoCheques", "Forms/ArchivoCheques.b1f")]
    class ArchivoCheques : UserFormBase
    {
        #region Propties
        bool allSelected;
        #endregion

        #region Constructor
        public ArchivoCheques()
        {
        }
        #endregion

        #region UI Components
        private SAPbouiCOM.Button Cancel;
        private SAPbouiCOM.Button btn_gena;
        private SAPbouiCOM.StaticText lbl_fdesde;
        private SAPbouiCOM.EditText txt_fdesde;
        private SAPbouiCOM.StaticText lbl_fhasta;
        private SAPbouiCOM.EditText txt_fhasta;
        private SAPbouiCOM.Button btn_find;
        private SAPbouiCOM.StaticText lbl_glosa;
        private SAPbouiCOM.EditText txt_glosa;
        private SAPbouiCOM.StaticText lbl_status;
        private SAPbouiCOM.ComboBox cbo_status;
        private SAPbouiCOM.Grid grd_checks;
        #endregion

        #region Initilizers
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Cancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.btn_gena = ((SAPbouiCOM.Button)(this.GetItem("btn_gena").Specific));
            this.btn_gena.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btn_gena_ClickAfter);
            this.lbl_fdesde = ((SAPbouiCOM.StaticText)(this.GetItem("lbl_fdesde").Specific));
            this.txt_fdesde = ((SAPbouiCOM.EditText)(this.GetItem("txt_fdesde").Specific));
            this.lbl_fhasta = ((SAPbouiCOM.StaticText)(this.GetItem("lbl_fhasta").Specific));
            this.txt_fhasta = ((SAPbouiCOM.EditText)(this.GetItem("txt_fhasta").Specific));
            this.btn_find = ((SAPbouiCOM.Button)(this.GetItem("btn_find").Specific));
            this.btn_find.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btn_find_ClickAfter);
            this.lbl_glosa = ((SAPbouiCOM.StaticText)(this.GetItem("lbl_glosa").Specific));
            this.txt_glosa = ((SAPbouiCOM.EditText)(this.GetItem("txt_glosa").Specific));
            this.lbl_status = ((SAPbouiCOM.StaticText)(this.GetItem("lbl_status").Specific));
            this.cbo_status = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_status").Specific));
            this.grd_checks = ((SAPbouiCOM.Grid)(this.GetItem("grd_checks").Specific));
            this.grd_checks.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.grd_checks_DoubleClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {
            this.UIAPIRawForm.DataSources.DBDataSources.Item("OADM").Query();
            this.UIAPIRawForm.DataSources.UserDataSources.Item("ud_fdesde").Value = DateTime.Now.ToString("yyyyMMdd");
            this.UIAPIRawForm.DataSources.UserDataSources.Item("ud_fhasta").Value = DateTime.Now.ToString("yyyyMMdd");
        }
        #endregion

        #region Events
        private void btn_find_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try 
            {
                allSelected = true;
                LoadGrid();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: ArchivoChequeCarga.Forms.ArchivoCheques.cs > btn_find_ClickAfter(): " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void grd_checks_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "Seleccionar":
                        if (pVal.Row == -1)
                        {
                            allSelected = !allSelected;
                            LoadGrid();
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: ArchivoChequeCarga.Forms.ArchivoCheques.cs > grd_checks_DoubleClickAfter(): " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void btn_gena_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            StringBuilder fileContent = new StringBuilder();
            XElement oXmlData = null;
            IEnumerable<XElement> oSelectedRows = null;
            try
            {
                if (this.UIAPIRawForm.DataSources.DataTables.Item("dt_checks").Rows.Count > 0)
                {
                    #region Variables
                    string MaxDate = "";
                    string nombreBeneficiario = "";
                    string rutBeneficiario = "";
                    string cuentaCorrienteCheque = "";
                    string numeroCheque = "";
                    int montoCheque = 0;
                    string marcaAnulacion = "";
                    string fechaCheque = "";
                    int totalRegistrosNoAnulados = 0;
                    int sumaRegistrosNoAnulados = 0;
                    int totalRegistrosAnulados = 0;
                    int sumaRegistrosAnulados = 0;
                    #endregion

                    #region Get Selected rows via XML
                    oXmlData = XElement.Parse(this.UIAPIRawForm.DataSources.DataTables.Item("dt_checks").SerializeAsXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly));
                    oSelectedRows = oXmlData.Descendants("Row").Where(row => row.Value.Contains("SeleccionarY"));
                    #endregion
                    if (oSelectedRows.Count() > 0)
                    {
                        #region Header
                        MaxDate = oSelectedRows.Descendants("Cell").Where(cell => cell.Value.Contains("Fecha Emision")).Descendants("Value").Min(x => int.Parse(x.Value)).ToString();
                        fileContent.AppendFormat("{0}", "1"); //Tipo Registro
                        fileContent.AppendFormat("{0}", this.UIAPIRawForm.DataSources.DBDataSources.Item("OADM").GetValue("PrintHeadr", 0).PadRight(35, ' ').Substring(0, 35));//Nombre Emisor
                        fileContent.AppendFormat("{0}", this.UIAPIRawForm.DataSources.DBDataSources.Item("OADM").GetValue("TaxIdNum", 0).Split('-')[0].PadLeft(8, '0').Substring(0, 8));//Rut Emisor
                        fileContent.AppendFormat("{0}", this.UIAPIRawForm.DataSources.DBDataSources.Item("OADM").GetValue("TaxIdNum", 0).Split('-')[1].PadLeft(1, '0').Substring(0, 1));//DV Emisor
                        fileContent.AppendFormat("{0}", this.txt_glosa.Value.PadRight(20, ' ').Substring(0, 20));//Glosa Archivo
                        fileContent.AppendFormat("{0}", DateTime.Now.ToString("yyyyMMdd"));//Fecha grabacion archivo
                        fileContent.AppendFormat("{0}", MaxDate.PadLeft(8, '0').Substring(0, 8));//Fecha cheque mas antiguo
                        fileContent.AppendFormat("{0}", "".PadRight(3, ' '));//Filler
                        fileContent.AppendFormat("{0}", "".PadRight(10, ' '));//Filler
                        fileContent.AppendFormat("{0}", "".PadRight(11, ' '));//Filler
                        fileContent.AppendLine();
                        #endregion

                        #region Detail
                        for (int i = 0; i < oSelectedRows.Count(); i++)
                        {
                            nombreBeneficiario = oSelectedRows.ElementAt(i).Descendants("Cell").Where(cell => cell.Value.Contains("Nombre. SN")).Descendants("Value").FirstOrDefault().Value;
                            rutBeneficiario = oSelectedRows.ElementAt(i).Descendants("Cell").Where(cell => cell.Value.Contains("RUT")).Descendants("Value").FirstOrDefault().Value;
                            cuentaCorrienteCheque = oSelectedRows.ElementAt(i).Descendants("Cell").Where(cell => cell.Value.Contains("Cuenta")).Descendants("Value").FirstOrDefault().Value;
                            numeroCheque = oSelectedRows.ElementAt(i).Descendants("Cell").Where(cell => cell.Value.Contains("N° Cheque")).Descendants("Value").FirstOrDefault().Value;
                            montoCheque = int.Parse(oSelectedRows.ElementAt(i).Descendants("Cell").Where(cell => cell.Value.Contains("Monto")).Descendants("Value").FirstOrDefault().Value.Split('.')[0]);
                            marcaAnulacion = oSelectedRows.ElementAt(i).Descendants("Cell").Where(cell => cell.Value.Contains("Anulado")).Descendants("Value").FirstOrDefault().Value;
                            fechaCheque = oSelectedRows.ElementAt(i).Descendants("Cell").Where(cell => cell.Value.Contains("Fecha Emision")).Descendants("Value").FirstOrDefault().Value;
                            fileContent.AppendFormat("{0}", "2"); //Tipo Registro
                            fileContent.AppendFormat("{0}", nombreBeneficiario.PadRight(50, ' ').Substring(0, 50)); //Nombre de Beneficiario
                            fileContent.AppendFormat("{0}", rutBeneficiario.Split('-')[0].PadLeft(8, '0').Substring(0, 8)); //Rut Beneficiario
                            fileContent.AppendFormat("{0}", rutBeneficiario.Split('-')[1].PadLeft(1, '0').Substring(0, 1)); //DV Beneficiario
                            fileContent.AppendFormat("{0}", cuentaCorrienteCheque.PadLeft(11, '0').Substring(0, 11)); //Cuenta Corriente Cheque
                            fileContent.AppendFormat("{0}", "".PadRight(2, ' ')); //Serie del Cheque
                            fileContent.AppendFormat("{0}", numeroCheque.PadLeft(7, '0').Substring(0, 7)); //Nro de cheque
                            fileContent.AppendFormat("{0}", montoCheque.ToString().PadLeft(13, '0').Substring(0, 13)); //Monto cheque
                            fileContent.AppendFormat("{0}", marcaAnulacion.PadRight(3, ' ').Substring(0, 3)); //Marca de Anulacion
                            fileContent.AppendFormat("{0}", fechaCheque.PadLeft(8, '0').Substring(0, 8)); //Tipo Registro
                            fileContent.AppendFormat("{0}", "".PadRight(1, ' '));//Filler
                            fileContent.AppendLine();
                            if (marcaAnulacion == "ANU")
                            {
                                totalRegistrosAnulados++;
                                sumaRegistrosAnulados += montoCheque;
                            }
                            else
                            {
                                totalRegistrosNoAnulados++;
                                sumaRegistrosNoAnulados += montoCheque;
                            }
                        }
                        #endregion

                        #region Totals
                        fileContent.AppendFormat("{0}", "3"); //Tipo Registro
                        fileContent.AppendFormat("{0}", totalRegistrosNoAnulados.ToString().PadLeft(10, '0').Substring(0, 10)); //Total Registros no anulados
                        fileContent.AppendFormat("{0}", sumaRegistrosNoAnulados.ToString().PadLeft(15, '0').Substring(0, 15)); //Suma Registros no anulados
                        fileContent.AppendFormat("{0}", totalRegistrosAnulados.ToString().PadLeft(10, '0').Substring(0, 10)); //Total Registros  anulados
                        fileContent.AppendFormat("{0}", sumaRegistrosAnulados.ToString().PadLeft(15, '0').Substring(0, 15)); //Suma Registros  anulados
                        fileContent.AppendFormat("{0}", (totalRegistrosNoAnulados + totalRegistrosAnulados).ToString().PadLeft(10, '0').Substring(0, 10)); //Total Registros 
                        fileContent.AppendFormat("{0}", (sumaRegistrosNoAnulados + sumaRegistrosAnulados).ToString().PadLeft(15, '0').Substring(0, 15)); //Suma Registros
                        fileContent.AppendFormat("{0}", "".PadRight(29, ' '));//Filler
                        #endregion

                        GuardarArchivo(fileContent.ToString());
                        this.grd_checks.DataTable.Clear();
                        this.UIAPIRawForm.DataSources.UserDataSources.Item("ud_fdesde").Value = DateTime.Now.ToString("yyyyMMdd");
                        this.UIAPIRawForm.DataSources.UserDataSources.Item("ud_fhasta").Value = DateTime.Now.ToString("yyyyMMdd");
                        this.UIAPIRawForm.DataSources.UserDataSources.Item("ud_glosa").Value ="";
                    }
                    else
                        Application.SBO_Application.MessageBox("No hay registros seleccionados para procesar");
                }
                else
                    Application.SBO_Application.MessageBox("No hay registros para procesar");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: ArchivoChequeCarga.Forms.ArchivoCheques.cs > btn_gena_ClickAfter(): " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        #endregion

        #region Methods
        private void LoadGrid()
        {
            this.UIAPIRawForm.Freeze(true);
            try
            {
                this.grd_checks.DataTable.ExecuteQuery(
                    Model.Consultas.ConsultaDetalleCheques(this.cbo_status.Selected.Value, this.txt_fdesde.Value, this.txt_fhasta.Value, allSelected)
                    );


                this.grd_checks.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

                this.grd_checks.Columns.Item(1).Editable = false;
                this.grd_checks.Columns.Item(1).TitleObject.Sortable = true;
                this.grd_checks.Columns.Item(1).Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                ((SAPbouiCOM.EditTextColumn)this.grd_checks.Columns.Item(1)).LinkedObjectType = "57";

                this.grd_checks.Columns.Item(2).Editable = false;
                this.grd_checks.Columns.Item(2).TitleObject.Sortable = true;
                this.grd_checks.Columns.Item(2).Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                ((SAPbouiCOM.EditTextColumn)this.grd_checks.Columns.Item(2)).LinkedObjectType = "2";

                this.grd_checks.Columns.Item(3).Editable = false;
                this.grd_checks.Columns.Item(3).TitleObject.Sortable = true;

                this.grd_checks.Columns.Item(4).Editable = false;
                this.grd_checks.Columns.Item(4).TitleObject.Sortable = true;

                this.grd_checks.Columns.Item(5).Editable = false;
                this.grd_checks.Columns.Item(5).TitleObject.Sortable = true;

                this.grd_checks.Columns.Item(6).Editable = false;
                this.grd_checks.Columns.Item(6).TitleObject.Sortable = true;

                this.grd_checks.Columns.Item(7).Editable = false;
                this.grd_checks.Columns.Item(7).TitleObject.Sortable = true;
                ((SAPbouiCOM.EditTextColumn)this.grd_checks.Columns.Item(7)).RightJustified = true;

                this.grd_checks.Columns.Item(8).Editable = false;
                this.grd_checks.Columns.Item(8).TitleObject.Sortable = true;

                this.grd_checks.Columns.Item(9).Editable = false;
                this.grd_checks.Columns.Item(9).TitleObject.Sortable = true;
              
                this.grd_checks.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: ArchivoChequeCarga.Forms.ArchivoCheques.cs > LoadGrid(): " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }
        }

        public void GuardarArchivo(string cadena)
        {
            string archivo = "";
            string path = "";

            try
            {
                archivo = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                path = archivo + @"\" + "ArchivoChequesCarga_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine(cadena);
                    sw.Close();
                }
                Application.SBO_Application.MessageBox("El archivo se genero de forma exitosa en: \n" + path);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: ArchivoChequeCarga.Forms.ArchivoCheques.cs > GuardarArchivo(): " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        } 
        #endregion
    }
}
