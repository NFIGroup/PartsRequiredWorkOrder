﻿using Parts_Required_from_WorkOrder.RightNowService;
using RightNow.AddIns.AddInViews;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Parts_Required_from_WorkOrder.WebServiceReqdParams;

namespace Parts_Required_from_WorkOrder
{
    class RightNowConnectService
    {
        private static RightNowConnectService _rightnowConnectService;
        private static object _sync = new object();
        private static RightNowSyncPortClient _rightNowClient;
     
        private RightNowConnectService()
        {

        }
        public static RightNowConnectService GetService()
        {
            if (_rightnowConnectService != null)
            {
                return _rightnowConnectService;
            }

            try
            {
                lock (_sync)
                {
                    if (_rightnowConnectService == null)
                    {
                        // Initialize client with current interface soap url 
                        string url = ReportCommandAddIn._globalContext.GetInterfaceServiceUrl(ConnectServiceType.Soap);
                        EndpointAddress endpoint = new EndpointAddress(url);

                        BasicHttpBinding binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportWithMessageCredential);
                        binding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;

                        // Optional depending upon use cases
                        binding.MaxReceivedMessageSize = 1024 * 1024;
                        binding.MaxBufferSize = 1024 * 1024;
                        binding.MessageEncoding = WSMessageEncoding.Mtom;

                        _rightNowClient = new RightNowSyncPortClient(binding, endpoint);

                        BindingElementCollection elements = _rightNowClient.Endpoint.Binding.CreateBindingElements();
                        elements.Find<SecurityBindingElement>().IncludeTimestamp = false;
                        _rightNowClient.Endpoint.Binding = new CustomBinding(elements);
                        ReportCommandAddIn._globalContext.PrepareConnectSession(_rightNowClient.ChannelFactory);

                        _rightnowConnectService = new RightNowConnectService();
                    }

                }
            }
            catch (Exception e)
            {
                _rightnowConnectService = null;
               
            }

            return _rightnowConnectService;
        }
        /// <summary>
        /// Return individual fields as per query
        /// </summary>
        /// <param name="ApplicationID"></param>
        /// <param name="Query"></param>
        /// <returns> array of string delimited by '~'</returns>
        private string[] GetRNData(string ApplicationID, string Query)
        {
            string[] rnData = null;
            ClientInfoHeader hdr = new ClientInfoHeader() { AppID = ApplicationID };

            byte[] output = null;
            CSVTableSet data = null;

            try
            {
                data = _rightNowClient.QueryCSV(hdr, Query, 1000, "~", false, false, out output);
                string dataRow = String.Empty;
                if (data != null && data.CSVTables.Length > 0 && data.CSVTables[0].Rows.Length > 0)
                {
                    return data.CSVTables[0].Rows;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception in GetRN: "+ex.Message);
            }
            return rnData;
        }

        /// <summary>
        /// Get ebs id of site
        /// </summary>
        /// <param name="siteID"></param>
        /// <returns> string </returns>
        public string GetEbsID(int siteID)
        {
            string query = String.Format("SELECT ebs_id_site FROM CO.Site WHERE ID = {0} limit 1", siteID);
            string[] resultSet = GetRNData("Get Ebs Site Id", query);
            if (resultSet != null && resultSet.Length > 0)
            {
                return resultSet[0];
            }
            return "";
        }
        /// <summary>
        /// Get Bus owner org ebs id
        /// </summary>
        /// <param name="vinID"></param>
        /// <returns> string </returns>
        public string GetCustomerEbsID(int orgID)
        {
            string query = string.Format("SELECT O.CustomFields.Co.ebs_id_org FROM Organization O WHERE O.ID = {0} limit 1", orgID);
            string[] resultSet = GetRNData(" Get Org Ebs ID", query);
            if (resultSet != null && resultSet.Length > 0)
            {
                return resultSet[0];
            }
            return "";
        }
        /// <summary>
        /// Get Reported Incident Ref num
        /// </summary>
        /// <param name="incVinID"></param>
        /// <returns> string </returns>
        public string GetRefNum(int incVinID)
        {
            string query = "SELECT IV.Incident.CustomFields.CO.reporting_incident.ReferenceNumber FROM "+
                           "CO.Incident_VIN as IV where ID = "+ incVinID;
            string[] resultSet = GetRNData(" Get Ref Num", query);
            if (resultSet != null && resultSet.Length > 0)
            {
                return resultSet[0];
            }
            return "";
        }
        /// <summary>
        /// Get Order type label
        /// </summary>
        /// <param name="orderTypeID"></param>
        /// <returns> order type label</returns>
        public string GetOrderTypeName(int orderTypeID)
        {
            string query = "SELECT Name FROM CO.OrderType WHERE ID = " + orderTypeID + " limit 1";
            string[] resultSet = GetRNData("Order Type Info", query);
            if (resultSet != null && resultSet.Length > 0)
            {
                return resultSet[0];
            }
            return "";
        }
        /// <summary>
        /// Get Distribution Center label
        /// </summary>
        /// <param name="distributionCenterID"></param>
        /// <returns> distribution Center label</returns>
        public string GetDistributionCenterIDName(int distributionCenterID)
        {
            string query = "SELECT Name FROM CO.Distribution_Center WHERE ID = " + distributionCenterID + " limit 1";
            string[] resultSet = GetRNData("distribution Center Info", query);
            if (resultSet != null && resultSet.Length > 0)
            {
                return resultSet[0];
            }
            return "";
        }
        /// <summary>
        /// Method which is called to get value of custom object.
        /// </summary>
        /// <param name="fieldName">The name of the custom field.</param>
        /// <returns>Value of the field</returns>
        public string GetCOFieldValue(string fieldName, IGenericObject customObject)
        {
            IList<IGenericField> fields = customObject.GenericFields;
            if (null != fields)
            {
                foreach (IGenericField field in fields)
                {
                    if (field.Name.Equals(fieldName))
                    {
                        if (field.DataValue.Value != null)
                            return field.DataValue.Value.ToString();
                    }
                }
            }
            return "";
        }
        /// <summary>
        /// Get Config Value based on lookupName
        /// </summary>
        /// <param name="configLookupName"></param>
        /// <returns>config value</returns>
        public string GetConfigValue(string configLookupName)
        {
            string query = "select Configuration.Value from Configuration where lookupname = '" + configLookupName + "'";
            string[] resultSet = GetRNData("Configuartion Value", query);
            if (resultSet != null && resultSet.Length > 0)
            {
                var jsonTrim = resultSet[0].Replace("\"\"", "\"");

                // jsonString has extra " at start, end and each " 
                int i = jsonTrim.IndexOf("\"");
                int j = jsonTrim.LastIndexOf("\"");
                var finalJson = jsonTrim.Substring(i + 1, j - 1);
                finalJson = finalJson.Replace("@xmlns", "xmlns");

                return finalJson;
            }
            return null;
        }

        /// <summary>
        /// To Create Parts Order Record
        /// </summary>
        /// <param name="partsID">Parent Parts ID</param>
        /// <returns></returns>
        public int CreatePartsOrder(int partsID)
        {
            try
            {
                GenericObject partsOrderObject = new GenericObject();
                partsOrderObject.ObjectType = new RNObjectType
                {
                    Namespace = "CO",
                    TypeName = "PartsOrder"
                };
                List<GenericField> genericFields = new List<GenericField>();
                genericFields.Add(createGenericField("Parts", createNamedIDDataValue(partsID), DataTypeEnum.NAMED_ID));
                partsOrderObject.GenericFields = genericFields.ToArray();

                ClientInfoHeader hdr = new ClientInfoHeader() { AppID = "Create Parts Order" };

                RNObject[] resultObj = _rightNowClient.Create(hdr, new RNObject[] { partsOrderObject }, new CreateProcessingOptions { SuppressExternalEvents = false, SuppressRules = false });
                if (resultObj != null)
                {
                    return Convert.ToInt32(resultObj[0].ID.id);
                }
            }
            catch (Exception ex)
            {
                ReportCommandAddIn.form.Hide();
                MessageBox.Show("Exception in creating Parts Order Record: " + ex.Message);
            }
            return 0;
        }
        /// <summary>
        /// To update PartsOrder object to keep log of info send to EBS for parts order
        /// </summary>
        /// <param name="lineRecords">Line record passed to EBS</param>
        /// <param name="salesOrderNo"></param>
        /// <param name="numOfVIN"></param>
        /// <param name="partsHeaderRecord"></param>
        /// <param name="shipToSiteID"></param>
        /// <param name="partOdrInstrID"></param>
        public void UpdatePartsOrder(List<OELINEREC> lineRecords, string salesOrderNo, POEHEADERREC partsHeaderRecord,
                                     int shipToSiteID)
        {
            try
            {
                List<RNObject> rnObjs = new List<RNObject>();
                //Loop over each partsIDToBeOrder that has parts ID that's order recently
                foreach (OELINEREC lineRecord in lineRecords)
                {
                    GenericObject genObject = new GenericObject();
                    genObject.ObjectType = new RNObjectType
                    {
                        Namespace = "CO",
                        TypeName = "PartsOrder"
                    };
                    genObject.ID = new ID { id = lineRecord.ORDERED_ID, idSpecified = true };

                    List<GenericField> customFields = new List<GenericField>();
                    customFields.Add(createGenericField("SalesOrderNumber", createStringDataValue(salesOrderNo), DataTypeEnum.STRING));
                    customFields.Add(createGenericField("OrderType", createStringDataValue(partsHeaderRecord.ORDER_TYPE), DataTypeEnum.STRING));
                    customFields.Add(createGenericField("RequestedShipDate", createDateDataValue(partsHeaderRecord.SHIP_DATE), DataTypeEnum.STRING));
                    customFields.Add(createGenericField("NumberOfVINS", createIntegerDataValue(1), DataTypeEnum.INTEGER));
                    //customFields.Add(createGenericField("SRNumber", createStringDataValue("SR-"+partsHeaderRecord.PROJECT_NUMBER.Substring(2)), DataTypeEnum.STRING));
                    customFields.Add(createGenericField("ShipToSite", createNamedIDDataValue(shipToSiteID), DataTypeEnum.NAMED_ID));

                    genObject.GenericFields = customFields.ToArray();
                    rnObjs.Add(genObject);
                }
                callBatchJob(getUpdateMsg(rnObjs));
            }
            catch (Exception ex)
            {
                ReportCommandAddIn.form.Hide();
                MessageBox.Show("Exception in updating PartsOrder Record: " + ex.Message);
            }
        }
        /// <summary>
        /// To Destroy Parts Order Record if EBS integration fails
        /// </summary>
        /// <returns></returns>
        public void DestroyPartsOrder(List<OELINEREC> lineRecords)
        {
            try
            {
                List<RNObject> rnObjs = new List<RNObject>();
                //Loop over each partsIDToBeOrder that has parts ID that's order recently
                foreach (OELINEREC lineRecord in lineRecords)
                {
                    GenericObject genObject = new GenericObject();
                    genObject.ObjectType = new RNObjectType
                    {
                        Namespace = "CO",
                        TypeName = "PartsOrder"
                    };
                    genObject.ID = new ID
                    {
                        id = lineRecord.ORDERED_ID,
                        idSpecified = true
                    };
                    rnObjs.Add(genObject);
                }
                callBatchJob(getDestroyMsg(rnObjs));
            }
            catch (Exception ex)
            {
                ReportCommandAddIn.form.Hide();
                MessageBox.Show("Exception in creating Parts Order Record: " + ex.Message);
            }
            return;
        }
 
        #region Miscellaneous

        /// <summary>
        /// Perform Batch operation
        /// </summary>
        /// <param name="partMsg"></param>
        /// <param name="laborMsg"></param>
        public void callBatchJob(Object msg)
        {
            try
            {
                /*** Form BatchRequestItem structure ********************/

                BatchRequestItem[] requestItems = new BatchRequestItem[1];

                BatchRequestItem requestItem = new BatchRequestItem();

                requestItem.Item = msg;

                requestItems[0] = requestItem;
                requestItems[0].CommitAfter = true;
                requestItems[0].CommitAfterSpecified = true;
                /*********************************************************/


                ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
                clientInfoHeader.AppID = "Batcher";

                BatchResponseItem[] batchRes = _rightNowClient.Batch(clientInfoHeader, requestItems);
                //If response type is RequestErrorFaultType then show the error msg 
                if (batchRes[0].Item.GetType().Name == "RequestErrorFaultType")
                {
                    RequestErrorFaultType requestErrorFault = (RequestErrorFaultType)batchRes[0].Item;
                    MessageBox.Show("There is an error with batch job :: " + requestErrorFault.exceptionMessage);
                }
            }
            catch (FaultException ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }
        /// <summary>
        /// Delete Message for Batch
        /// </summary>
        /// <param name="coList"></param>
        /// <returns></returns>
        private DestroyMsg getDestroyMsg(List<RNObject> coList)
        {
            DestroyMsg destroyMsg = new DestroyMsg();
            DestroyProcessingOptions destroyProcessingOptions = new DestroyProcessingOptions();
            destroyProcessingOptions.SuppressExternalEvents = true;
            destroyProcessingOptions.SuppressRules = true;
            destroyMsg.ProcessingOptions = destroyProcessingOptions;

            destroyMsg.RNObjects = coList.ToArray();

            return destroyMsg;
        }
        /// <summary>
        /// To create Update Message for Batch
        /// </summary>
        /// <param name="coList"></param>
        /// <returns></returns>
        private UpdateMsg getUpdateMsg(List<RNObject> coList)
        {
            UpdateMsg updateMsg = new UpdateMsg();
            UpdateProcessingOptions updateProcessingOptions = new UpdateProcessingOptions();
            updateProcessingOptions.SuppressExternalEvents = true;
            updateProcessingOptions.SuppressRules = true;
            updateMsg.ProcessingOptions = updateProcessingOptions;

            updateMsg.RNObjects = coList.ToArray();

            return updateMsg;
        }

        /// <summary>
        /// To create Create Message for Batch
        /// </summary>
        /// <param name="coList"></param>
        /// <returns></returns>
        private CreateMsg getCreateMsg(List<RNObject> coList)
        {
            CreateMsg createMsg = new CreateMsg();
            CreateProcessingOptions createProcessingOptions = new CreateProcessingOptions();
            createProcessingOptions.SuppressExternalEvents = true;
            createProcessingOptions.SuppressRules = true;
            createMsg.ProcessingOptions = createProcessingOptions;

            createMsg.RNObjects = coList.ToArray();

            return createMsg;
        }

        /// <summary>
        /// Create string type data value
        /// </summary>
        /// <param name="val"></param>
        /// <returns> DataValue</returns>
        private DataValue createStringDataValue(string val)
        {
            DataValue dv = new DataValue();
            dv.Items = new Object[] { val };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.StringValue };  //Change this to the type of field
            return dv;
        }
        /// <summary>
        /// Create string type data value
        /// </summary>
        /// <param name="val"></param>
        /// <returns> DataValue</returns>
        private DataValue createDateDataValue(string val)
        {
            DateTime date = (Convert.ToDateTime(val));
            DataValue dv = new DataValue();
            dv.Items = new Object[] { date };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.DateValue };  //Change this to the type of field
            return dv;
        }
        /// <summary>
        /// Create Boolean type data value
        /// </summary>
        /// <param name="val"></param>
        /// <returns> DataValue</returns>
        private DataValue createBooleanDataValue(Boolean val)
        {
            DataValue dv = new DataValue();
            dv.Items = new Object[] { val };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.BooleanValue };

            return dv;
        }

        /// <summary>
        /// Create integer type data value
        /// </summary>
        /// <param name="val"></param>
        /// <returns> DataValue</returns>
        private DataValue createIntegerDataValue(int val)
        {
            DataValue dv = new DataValue();
            dv.Items = new Object[] { val };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.IntegerValue };  //Change this to the type of field
            return dv;
        }

        /// <summary>
        /// Create GenericField object
        /// </summary>
        /// <param name="name">Name Of Generic Field</param>
        /// <param name="dataValue">Vlaue of generic field</param>
        /// <param name="type">Type of generic field</param>
        /// <returns> GenericField</returns>
        private GenericField createGenericField(string name, DataValue dataValue, DataTypeEnum type)
        {
            GenericField genericField = new GenericField();

            genericField.dataType = type;
            genericField.dataTypeSpecified = true;
            genericField.name = name;
            genericField.DataValue = dataValue;
            return genericField;
        }

        /// <summary>
        /// Create Named ID type Data Value for NamedID as input
        /// </summary>
        /// <param name="namedvalue"></param>
        /// <returns></returns>
        private DataValue createNamedID(NamedID namedvalue)
        {
            DataValue dv = new DataValue();
            dv.Items = new Object[] { namedvalue };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.NamedIDValue };
            return dv;
        }

        /// <summary>
        /// Create Named ID type data value for integer value as input
        /// </summary>
        /// <param name="idVal"></param>
        /// <returns> DataValue</returns>
        private DataValue createNamedIDDataValue(long idVal)
        {
            ID id = new ID();
            id.id = idVal;
            id.idSpecified = true;

            NamedID namedID = new NamedID();
            namedID.ID = id;

            DataValue dv = new DataValue();
            dv.Items = new Object[] { namedID };
            dv.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.NamedIDValue };

            return dv;
        }
        #endregion
    }
}
