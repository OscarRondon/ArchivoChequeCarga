using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace ArchivoChequeCarga.Model
{
    class Consultas
    {
        #region Methods
        public static string ConsultaDetalleCheques(string status, string fdesde, string fhasta, bool allSelected)
        {
            StringBuilder  resp = new StringBuilder();
            try
            {
                resp.Append("select ");
                if(allSelected)
                    resp.Append("'Y' as \"Seleccionar\", ");
                else
                    resp.Append("'N' as \"Seleccionar\", ");
                resp.Append("T0.\"CheckKey\" as \"Id Cheque\", ");
                resp.Append("T1.\"CardCode\" as \"Cod. SN\", ");
                resp.Append("T1.\"CardName\" as \"Nombre. SN\", ");
                resp.Append("T1.\"LicTradNum\" as \"RUT\", ");
                resp.Append("T0.\"AcctNum\" as \"Cuenta\", ");
                resp.Append("T0.\"CheckNum\" as \"N° Cheque\", ");
                resp.Append("T0.\"CheckSum\" as \"Monto\", ");
                resp.Append("(CASE WHEN T0.\"Canceled\" = 'Y' THEN 'ANU' ELSE '' END) as \"Anulado\", ");
                resp.Append("T0.\"PmntDate\" as \"Fecha Emision\" ");
                resp.Append("from \"OCHO\" T0 ");
                resp.Append("inner join \"OCRD\" T1 on T0.\"VendorCode\" = T1.\"CardCode\" ");
                resp.Append("where ");
                resp.AppendFormat("T0.\"PmntDate\" between '{0}' and '{1}' ", fdesde, fhasta);
                if (!string.IsNullOrEmpty(status))
                    resp.AppendFormat("and T0.\"Canceled\" = '{0}' ", status);
                resp.Append("order by T0.\"CheckKey\" desc");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("Error: ArchivoChequeCarga.Model.Consultas.cs > ConsultaDetalleCheques(): " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return resp.ToString();
        }
        #endregion
    }
}
