using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;
using static Parts_Required_from_WorkOrder.WebServiceReqdParams;

namespace Parts_Required_from_WorkOrder
{
    [AddIn("Report Command AddIn", Version = "1.0.0.0")]
    public class ReportCommandAddIn : IReportCommand2
    {
        static public IGlobalContext _globalContext;
        IRecordContext _recordContext;
        IList<IReportRow> _selectedRows;
        IGenericObject _workOrder;
        IGenericObject _workOrderExtra;
        public static ProgressForm form = new ProgressForm();

        #region IReportCommand Members

        /// <summary>
        /// 
        /// </summary>
        public bool Enabled(IList<IReportRow> rows)
        {
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Execute(IList<IReportRow> rows)
        {
            _selectedRows = rows;
            _recordContext = _globalContext.AutomationContext.CurrentWorkspace;
            _workOrder = (IGenericObject)_recordContext.GetWorkspaceRecord("TOA$Work_Order");
            _workOrderExtra = (IGenericObject)_recordContext.GetWorkspaceRecord("TOA$Work_Order_Extra");

            System.Threading.Thread th = new System.Threading.Thread(ProcessSelectedRowInfo);
            th.Start();
            form.Show();
        }

        /// <summary>
        /// 
        /// </summary>
        public Image Image16
        {
            get
            {
                return Properties.Resources.AddIn16;
            }
        }

        /// <summary>
        /// To get Parts details and Call Webservice
        /// </summary>
        public void ProcessSelectedRowInfo()
        {
            POEHEADERREC partsHeaderRecord = new POEHEADERREC();
            RightNowConnectService _rnConnectService = RightNowConnectService.GetService();
            /*Get Work Order Info*/

            //Get ship set from Current Reported Incident
            string shipSet = _rnConnectService.GetCOFieldValue("ship_set", _workOrder);

            //Get Num of VIN for which this order will be placed
            int numOfVIN = 1;//will always be one

            string orderTypeId = _rnConnectService.GetCOFieldValue("Order_Type", _workOrderExtra);
            if (orderTypeId == String.Empty)
            {
                form.Hide();
                MessageBox.Show("Order Type is empty");
                return;
            }
            partsHeaderRecord.ORDER_TYPE = _rnConnectService.GetOrderTypeName(Convert.ToInt32(orderTypeId));

            string shipToId = _rnConnectService.GetCOFieldValue("Ship_To_Site", _workOrder);
            if (shipToId == String.Empty)
            {
                form.Hide();
                MessageBox.Show("Ship to Site is empty");
                return;
            }
            partsHeaderRecord.SHIP_TO_ORG_ID = Convert.ToInt32(_rnConnectService.GetEbsID(Convert.ToInt32(shipToId)));

            string shipDate = _rnConnectService.GetCOFieldValue("requested_ship_date", _workOrder);
            if (shipDate == String.Empty)
            {
                shipDate = DateTime.Today.ToString();//set to today's date
            }
            partsHeaderRecord.SHIP_DATE = Convert.ToDateTime(shipDate).ToString("dd-MMM-yyyy");

            string dcID = _rnConnectService.GetCOFieldValue("Distribution_center", _workOrderExtra);
            if (dcID == String.Empty)
            {
                form.Hide();
                MessageBox.Show("Distribution Center is empty");
                return;
            }
            partsHeaderRecord.WAREHOUSE_ORG = _rnConnectService.GetDistributionCenterIDName(Convert.ToInt32(dcID));

            string Organization = _rnConnectService.GetCOFieldValue("Organization", _workOrder);            
            partsHeaderRecord.CUSTOMER_ID = Convert.ToInt32(_rnConnectService.GetCustomerEbsID(Convert.ToInt32(Organization)));

            string incVinID = _rnConnectService.GetCOFieldValue("Incident_VIN_ID", _workOrder); 
            partsHeaderRecord.CLAIM_NUMBER = _rnConnectService.GetRefNum(Convert.ToInt32(incVinID));

            partsHeaderRecord.PROJECT_NUMBER = _rnConnectService.GetCOFieldValue("Project_Number", _workOrder);
           // partsHeaderRecord.RETROFIT_NUMBER = _rnConnectService.GetCOFieldValue("retrofit_number", _workOrder);
            partsHeaderRecord.CUST_PO_NUMBER = _rnConnectService.GetCOFieldValue("PO_Number", _workOrder);//
            partsHeaderRecord.SHIPPING_INSTRUCTIONS = _rnConnectService.GetCOFieldValue("Shipping_Instruction", _workOrder);

            List<OELINEREC> lineRecords = new List<OELINEREC>();

            //Loop over each selected parts that user wants to order
            //Get required info needed for EBS web-service
            foreach (IReportRow row in _selectedRows)
            {
                OELINEREC lineRecord = new OELINEREC();
                IList<IReportCell> cells = row.Cells;

                foreach (IReportCell cell in cells)
                {
                    if (cell.Name == "Part ID")
                    {
                        //Create PartsOrder record first as EBS web-service need PartsOrder ID
                        lineRecord.ORDERED_ID = _rnConnectService.CreatePartsOrder(Convert.ToInt32(cell.Value));//Pass parts Order ID
                    }
                    if (cell.Name == "Part #")
                    {
                        lineRecord.ORDERED_ITEM = cell.Value;
                    }
                    if (cell.Name == "Qty")
                    {
                        lineRecord.ORDERED_QUANTITY = Convert.ToDouble(Math.Round(Convert.ToDecimal(cell.Value), 2).ToString()) * numOfVIN;//get total qyantity 
                    }
                    if (cell.Name == "Source Type")
                    {
                        lineRecord.SOURCE_TYPE = cell.Value;
                    }
                }
                lineRecord.SHIP_SET = shipSet;
                lineRecords.Add(lineRecord);
            }
            //Call PartsRequired Model to send parts info to EBS
            PartsRequiredModel partsRequired = new PartsRequiredModel(_recordContext);
            partsRequired.GetDetails(lineRecords, _workOrder, partsHeaderRecord, Convert.ToInt32(shipToId));
            form.Hide();
            _recordContext.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Save);
        }

        /// <summary>
        /// 
        /// </summary>
        public Image Image32
        {
            get
            {
                return Properties.Resources.AddIn32;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public IList<RightNow.AddIns.Common.ReportRecordIdType> RecordTypes
        {
            get
            {
                IList<ReportRecordIdType> typeList = new List<ReportRecordIdType>();

                typeList.Add(ReportRecordIdType.Answer);
                typeList.Add(ReportRecordIdType.Chat);
                typeList.Add(ReportRecordIdType.CloudAcct2Search);
                typeList.Add(ReportRecordIdType.Contact);
                typeList.Add(ReportRecordIdType.ContactList);
                typeList.Add(ReportRecordIdType.Document);
                typeList.Add(ReportRecordIdType.Flow);
                typeList.Add(ReportRecordIdType.Incident);
                typeList.Add(ReportRecordIdType.Mailing);
                typeList.Add(ReportRecordIdType.MetaAnswer);
                typeList.Add(ReportRecordIdType.Opportunity);
                typeList.Add(ReportRecordIdType.Organization);
                typeList.Add(ReportRecordIdType.Question);
                typeList.Add(ReportRecordIdType.QueuedReport);
                typeList.Add(ReportRecordIdType.Quote);
                typeList.Add(ReportRecordIdType.QuoteProduct);
                typeList.Add(ReportRecordIdType.Report);
                typeList.Add(ReportRecordIdType.Segment);
                typeList.Add(ReportRecordIdType.Survey);
                typeList.Add(ReportRecordIdType.Task);
                typeList.Add(ReportRecordIdType.CustomObjectAll);

                return typeList;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Text
        {
            get
            {
                return "Order Parts ";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "Call Parts Required Webservice to send parts order";
            }
        }

        public IList<string> CustomObjectRecordTypes
        {
            get
            {
                IList<string> typeList = new List<string>();

                typeList.Add("PartsOrder");
                typeList.Add("Parts_PartsOrder");
                typeList.Add("Parts");
                typeList.Add("Work_Order");
                return typeList;
            }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            _globalContext = GlobalContext;
            return true;
        }

        #endregion
    }
}
