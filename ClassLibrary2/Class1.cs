using LSEXT;
using LSSERVICEPROVIDERLib;
using Oracle.ManagedDataAccess.Client;
using Patholab_Common;
using Patholab_DAL_V1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ClassLibrary2
{
    [ComVisible(true)]
    [ProgId("SendToPacs")]
    public class Class1 : IWorkflowExtension
    {
        private double sessionId;
        private string _connectionString;

        private OracleConnection _connection;

        INautilusServiceProvider sp;
        private DataLayer dal;

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
                string tableName = Parameters["TABLE_NAME"];
                string role = Parameters["ROLE_NAME"];

                sp = Parameters["SERVICE_PROVIDER"];
                var rs = Parameters["RECORDS"];


                ////////////יוצר קונקשן//////////////////////////
                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);
                /////////////////////////////           

                _connection = GetConnection(ntlCon);

                //     MessageBox.Show("0.1.Before new DataLayer");

                dal = new DataLayer();
                dal.Connect(ntlCon);

                string fullPath = "C:\\Temp\\";

                if (tableName == "ALIQUOT")
                {
                    rs.MoveLast();
                    double aliquotId = rs.Fields["ALIQUOT_ID"].Value;
                    //Creating the full path of the xml
                    fullPath += aliquotId.ToString() + ".xml";
                    ALIQUOT aliquot = dal.FindBy<ALIQUOT>(x => x.ALIQUOT_ID == aliquotId).FirstOrDefault();
                    ALIQUOT_USER aliquotUser = null;
                    if (aliquot != null)
                    {
                        aliquotUser = dal.FindBy<ALIQUOT_USER>(x => x.ALIQUOT_ID == aliquotId).FirstOrDefault();
                    }
                    SAMPLE sample = dal.FindBy<SAMPLE>(x => x.SAMPLE_ID == aliquot.SAMPLE_ID).FirstOrDefault();
                    SAMPLE_USER sampleUser = null;
                    if (sample != null)
                    {
                        sampleUser = dal.FindBy<SAMPLE_USER>(x => x.SAMPLE_ID == sample.SAMPLE_ID).FirstOrDefault();
                    }
                    SDG sdg = dal.FindBy<SDG>(x => x.SDG_ID == sample.SDG_ID).FirstOrDefault();
                    SDG_USER sdgUser = null;
                    if (sdg != null)
                    {
                        sdgUser = dal.FindBy<SDG_USER>(x => x.SDG_ID == sdg.SDG_ID).FirstOrDefault();
                    }
                    CLIENT client = dal.FindBy<CLIENT>(x => x.CLIENT_ID == sample.CLIENT_ID).FirstOrDefault();
                    CLIENT_USER clientUser=null;
                    if (client != null)
                    {
                        clientUser = dal.FindBy<CLIENT_USER>(x => x.CLIENT_ID == client.CLIENT_ID).FirstOrDefault();
                    }
                    
                    U_ORDER order = dal.FindBy<U_ORDER>(x => x.U_ORDER_ID == sdgUser.U_ORDER_ID).FirstOrDefault();
                    U_ORDER_USER orderUser = null;
                    if (order != null)
                    {
                        orderUser = dal.FindBy<U_ORDER_USER>(x => x.U_ORDER_ID == order.U_ORDER_ID).FirstOrDefault();
                    }
                    TEST test = dal.FindBy<TEST>(x => x.ALIQUOT_ID == aliquot.ALIQUOT_ID).FirstOrDefault();
                    RESULT result = dal.FindBy<RESULT>(x => x.TEST_ID == test.TEST_ID).FirstOrDefault();
                    //
                    SUPPLIER supplier = dal.FindBy<SUPPLIER>(x => x.SUPPLIER_ID == aliquot.SUPPLIER_ID).FirstOrDefault();
                    SUPPLIER_USER supplierUser = null;
                    if (supplier != null)
                    {
                        supplierUser = dal.FindBy<SUPPLIER_USER>(x => x.SUPPLIER_ID == supplier.SUPPLIER_ID).FirstOrDefault();
                    }
                    
                    U_CONTAINER container = dal.FindBy<U_CONTAINER>(x => x.U_CONTAINER_ID == sdgUser.U_CONTAINER_ID).FirstOrDefault();
                    U_CONTAINER_USER containerUser = null;
                    if (container != null)
                    {
                        containerUser = dal.FindBy<U_CONTAINER_USER>(x => x.U_CONTAINER_ID == container.U_CONTAINER_ID).FirstOrDefault();
                    }
                    U_DEBIT debit = dal.FindBy<U_DEBIT>(x => x.U_DEBIT_ID == aliquotUser.U_DEBIT_ID).FirstOrDefault();
                    U_DEBIT_USER debitUser = null;
                    if (debit != null)
                    {
                        debitUser = dal.FindBy<U_DEBIT_USER>(x => x.U_DEBIT_ID == debit.U_DEBIT_ID).FirstOrDefault();
                    }
                    U_CLINIC clinic = sdgUser.IMPLEMENTING_CLINIC;
                    U_CLINIC_USER clinicUser = null;
                    if (clinic != null)
                    {
                        clinicUser = dal.FindBy<U_CLINIC_USER>(x => x.U_CLINIC_ID == clinic.U_CLINIC_ID).FirstOrDefault();

                    }
                    U_CUSTOMER customer = dal.FindBy<U_CUSTOMER>(x => x.U_CUSTOMER_ID == clinicUser.U_CUSTOMER_ID).FirstOrDefault();
                    U_CUSTOMER_USER customerUser = null;
                    if (customer != null)
                    {
                        customerUser = dal.FindBy<U_CUSTOMER_USER>(x => x.U_CUSTOMER_ID == customer.U_CUSTOMER_ID).FirstOrDefault();
                    }

                    if(aliquotUser.U_GLASS_TYPE == "S"){
                        using (XmlWriter writer = XmlWriter.Create(fullPath))
                        {
                            writer.WriteStartDocument();
                            //Must have a starter element that wraps the whole file.
                            writer.WriteStartElement("FILE_PACS");

                            writer.WriteStartElement("CLIENT");
                            writer.WriteElementString("NAME", client != null ? client.NAME : "");
                            writer.WriteElementString("Cleint_Id", client != null ? client.CLIENT_ID.ToString() : "");
                            writer.WriteElementString("Passport", "");//שדה לא קיים
                            writer.WriteElementString("group_id", client != null ? client.GROUP_ID.ToString() : "");
                            writer.WriteElementString("version", client != null ? client.VERSION : "");
                            writer.WriteElementString("description", client != null ? client.DESCRIPTION : "");
                            writer.WriteElementString("bad_debt", client != null ? client.BAD_DEBT : "");
                            writer.WriteElementString("version_status", client != null ? client.VERSION_STATUS : "");
                            writer.WriteElementString("Client_discount", client != null ? client.CLIENT_DISCOUNT.ToString() : "");
                            writer.WriteElementString("Parent_version_id", client != null ? client.PARENT_VERSION_ID.ToString() : "");
                            writer.WriteElementString("Client_type", client != null ? client.CLIENT_TYPE : "");
                            writer.WriteElementString("Client_code", client != null ? client.CLIENT_CODE : "");                          
                            writer.WriteEndElement();

                            writer.WriteStartElement("CLIENT_USER");
                            writer.WriteElementString("Cleint_Id",clientUser != null ? clientUser.CLIENT_ID.ToString() : "");
                            writer.WriteElementString("u_First_Name", clientUser != null ? clientUser.U_FIRST_NAME : "");
                            writer.WriteElementString("u_Last_Name", clientUser != null ? clientUser.U_LAST_NAME : "");
                            writer.WriteElementString("u_Date_Of_Birth", clientUser != null ? clientUser.U_DATE_OF_BIRTH.ToString() : "");
                            writer.WriteElementString("u_Phone", clientUser != null ? clientUser.U_PHONE : "");
                            writer.WriteElementString("u_Gender", clientUser != null ? clientUser.U_GENDER : "");
                            writer.WriteElementString("u_Age", clientUser != null ? clientUser.U_AGE.ToString() : "");
                            writer.WriteElementString("u_id_code", clientUser != null ? clientUser.U_ID_CODE : "");
                            writer.WriteElementString("u_passport", clientUser != null ? clientUser.U_PASSPORT : "");
                            writer.WriteElementString("U_phon_2", clientUser != null ? clientUser.U_PHON_2 : "");
                            writer.WriteElementString("u_visit_1", clientUser != null ? clientUser.U_VISIT_1 : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("SDG");
                            writer.WriteElementString("SDG_ID",sdg.SDG_ID.ToString());
                            writer.WriteElementString("GROUP_ID",sdg.GROUP_ID.ToString());
                            writer.WriteElementString("DELIVERY_DATE",sdg.DELIVERY_DATE.ToString());
                            writer.WriteElementString("ARCHIVED_CHILD_COMPLETE",sdg.ARCHIVED_CHILD_COMPLETE);
                            writer.WriteElementString("SDG_TEMPLATE_ID",sdg.SDG_TEMPLATE_ID.ToString());
                            writer.WriteElementString("WORKFLOW_NODE_ID",sdg.WORKFLOW_NODE_ID.ToString());
                             writer.WriteElementString("NAME",sdg.NAME);
                            writer.WriteElementString("DESCRIPTION",sdg.DESCRIPTION);
                            writer.WriteElementString("STATUS",sdg.STATUS);
                            writer.WriteElementString("OLD_STATUS",sdg.OLD_STATUS);
                            writer.WriteElementString("EVENTS",sdg.EVENTS);
                            writer.WriteElementString("INSPECTION_PLAN_ID",sdg.INSPECTION_PLAN_ID.ToString());
                             writer.WriteElementString("NEEDS_REVIEW",sdg.NEEDS_REVIEW);
                            writer.WriteElementString("RECEIVED_BY",sdg.RECEIVED_BY.ToString());
                            writer.WriteElementString("RECEIVED_ON",sdg.RECEIVED_ON.ToString());
                            writer.WriteElementString("AUTHORISED_ON",sdg.AUTHORISED_ON.ToString());
                            writer.WriteElementString("REPORTED",sdg.REPORTED);
                            writer.WriteElementString("HAS_NOTES",sdg.HAS_NOTES);
                             writer.WriteElementString("CONCLUSION",sdg.CONCLUSION);
                            writer.WriteElementString("HAS_AUDITS",sdg.HAS_AUDITS);
                            writer.WriteElementString("CREATED_ON",sdg.CREATED_ON.ToString());
                            writer.WriteElementString("AUTHORISED_BY",sdg.AUTHORISED_BY.ToString());
                            writer.WriteElementString("COMPLETED_ON",sdg.COMPLETED_ON.ToString());
                            writer.WriteElementString("EXTERNAL_REFERENCE",sdg.EXTERNAL_REFERENCE);
                            writer.WriteEndElement();

                            writer.WriteStartElement("SDG_USER");
                            writer.WriteElementString("SDG_ID",sdgUser.SDG_ID.ToString());
                            writer.WriteElementString("U_CLINIC_CODE",sdgUser.U_CLINIC_CODE);
                            writer.WriteElementString("U_CONSULT",sdgUser.U_CONSULT);
                            writer.WriteElementString("U_ISCONSULT",sdgUser.U_ISCONSULT);
                            writer.WriteElementString("U_ISPOSITIVE",sdgUser.U_ISPOSITIVE);
                            writer.WriteElementString("U_LAST_UPDATE",sdgUser.U_LAST_UPDATE.ToString());
                            writer.WriteElementString("U_MALIGNANT",sdgUser.U_MALIGNANT);
                            writer.WriteElementString("U_NBR_OF_SAMPLES",sdgUser.U_NBR_OF_SAMPLES.ToString());
                            writer.WriteElementString("U_PATHOLOG",sdgUser.U_PATHOLOG.ToString());
                            writer.WriteElementString("U_PATIENT",sdgUser.U_PATIENT.ToString());
                            writer.WriteElementString("U_PRIORITY",sdgUser.U_PRIORITY.ToString());
                            writer.WriteElementString("U_PREGNANCY",sdgUser.U_PREGNANCY);
                            writer.WriteElementString("U_PREGNANCY_DATE",sdgUser.U_PREGNANCY_DATE.ToString());
                            writer.WriteElementString("U_REFERRAL_DATE",sdgUser.U_REFERRAL_DATE.ToString());
                            writer.WriteElementString("U_REFERRING_PHYSICIAN",sdgUser.U_REFERRING_PHYSICIAN.ToString());
                            writer.WriteElementString("U_REPORTED",sdgUser.U_REPORTED);
                            writer.WriteElementString("U_REQUEST_DATE",sdgUser.U_REQUEST_DATE.ToString());
                            writer.WriteElementString("U_REVISION_CAUSE",sdgUser.U_REVISION_CAUSE);
                            writer.WriteElementString("U_SLIDE_NBR",sdgUser.U_SLIDE_NBR.ToString());
                            writer.WriteElementString("U_SNOMED",sdgUser.U_SNOMED);
                            writer.WriteElementString("U_SNOMED_T",sdgUser.U_SNOMED_T);
                            writer.WriteElementString("U_STORED",sdgUser.U_STORED);
                            writer.WriteElementString("U_SUSPENSION_CAUSE",sdgUser.U_SUSPENSION_CAUSE);
                            writer.WriteElementString("U_WEEK_NBR",sdgUser.U_WEEK_NBR.ToString());
                            writer.WriteElementString("U_IMPLEMENTING_PHYSICIAN",sdgUser.U_IMPLEMENTING_PHYSICIAN.ToString());
                            writer.WriteElementString("U_IMPLEMENTING_CLINIC",sdgUser.U_IMPLEMENTING_CLINIC.ToString());
                            writer.WriteElementString("U_COLLECTION_STATION",sdgUser.U_COLLECTION_STATION.ToString());
                            writer.WriteElementString("U_IS_QC",sdgUser.U_IS_QC);
                            writer.WriteElementString("U_ATFILENM",sdgUser.U_ATFILENM);
                            writer.WriteElementString("U_PATHOLAB_NUMBER",sdgUser.U_PATHOLAB_NUMBER.ToString());
                            writer.WriteElementString("U_HOSPITAL_NUMBER",sdgUser.U_HOSPITAL_NUMBER.ToString());
                            writer.WriteElementString("U_NO_OBLIGATION",sdgUser.U_NO_OBLIGATION);
                            writer.WriteElementString("U_CONTAINER_ID",sdgUser.U_CONTAINER_ID.ToString());
                            writer.WriteElementString("U_AGE_AT_ARRIVAL",sdgUser.U_AGE_AT_ARRIVAL.ToString());
                            writer.WriteElementString("U_RECEIVE_QC",sdgUser.U_RECEIVE_QC);
                            writer.WriteElementString("U_REFERRAL_PHYSICIAN_CLINI",sdgUser.U_REFERRAL_PHYSICIAN_CLINI.ToString());
                            writer.WriteElementString("U_PASSPORT",sdgUser.U_PASSPORT);
                            writer.WriteElementString("U_ORDER_ID",sdgUser.U_ORDER_ID.ToString());
                            writer.WriteElementString("U_ACCORD_NUMBER",sdgUser.U_ACCORD_NUMBER);
                            writer.WriteElementString("U_PDF_PATH",sdgUser.U_PDF_PATH);
                            writer.WriteElementString("U_FAX_EMAIL_SENT_ON",sdgUser.U_FAX_EMAIL_SENT_ON.ToString());
                            writer.WriteElementString("U_IS_LAST_UPDATE",sdgUser.U_IS_LAST_UPDATE);
                            writer.WriteElementString("U_IS_THINPREP",sdgUser.U_IS_THINPREP);
                            writer.WriteElementString("U_CALCULATE_DEBIT",sdgUser.U_CALCULATE_DEBIT);
                            writer.WriteElementString("U_TUMOR_SIZE",sdgUser.U_TUMOR_SIZE);
                            writer.WriteElementString("U_CANCELD_ON",sdgUser.U_CANCELD_ON.ToString());
                            writer.WriteElementString("U_CANCELD_BY",sdgUser.U_CANCELD_BY.ToString());
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_ORDER");
                            writer.WriteElementString("U_ORDER_ID",order != null ? order.U_ORDER_ID.ToString() : "");
                            writer.WriteElementString("NAME", order != null ? order.NAME : "");
                            writer.WriteElementString("DESCRIPTION", order != null ? order.DESCRIPTION : "");
                            writer.WriteElementString("VERSION", order != null ? order.VERSION : "");
                            writer.WriteElementString("VERSION_STATUS", order != null ? order.VERSION_STATUS : "");
                            writer.WriteElementString("GROUP_ID", order != null ? order.GROUP_ID.ToString() : "");
                            writer.WriteElementString("PARENT_VERSION_ID", order != null ? order.PARENT_VERSION_ID.ToString() : "");
                            writer.WriteElementString("TEMPLATE_ID", order != null ? order.TEMPLATE_ID.ToString() : "");
                            writer.WriteElementString("WORKFLOW_NODE_ID", order != null ? order.WORKFLOW_NODE_ID.ToString() : "");
                            writer.WriteElementString("EVENTS",order != null ? order.EVENTS : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_ORDER_USER");
                            writer.WriteElementString("U_STATUS", orderUser != null ? orderUser.U_STATUS : "");
                            writer.WriteElementString("U_INC_VAT", orderUser != null ? orderUser.U_INC_VAT : "");
                            writer.WriteElementString("U_TOTAL_DEBIT", orderUser != null ? orderUser.U_TOTAL_DEBIT.ToString() : "");
                            writer.WriteElementString("U_IN_ADVANCE", orderUser != null ? orderUser.U_IN_ADVANCE : "");
                            writer.WriteElementString("U_SDG_NAME", orderUser != null ? orderUser.U_SDG_NAME : "");
                            writer.WriteElementString("U_CUSTOMER", orderUser != null ? orderUser.U_CUSTOMER.ToString() : "");
                            writer.WriteElementString("U_ORDER_ID", orderUser != null ? orderUser.U_ORDER_ID.ToString() : "");
                            writer.WriteElementString("U_SECOND_CUSTOMER", orderUser != null ? orderUser.U_SECOND_CUSTOMER.ToString() : "");
                            writer.WriteElementString("U_URGENT", orderUser != null ? orderUser.U_URGENT : "");
                            writer.WriteElementString("U_PARTS_ID", orderUser != null ? orderUser.U_PARTS_ID.ToString() : "");
                            writer.WriteElementString("U_PAY_TYPE", orderUser != null ? orderUser.U_PAY_TYPE : "");
                            writer.WriteElementString("U_PAY_AMOUNT", orderUser != null ? orderUser.U_PAY_AMOUNT.ToString() : "");
                            writer.WriteElementString("U_CREATED_ON", orderUser != null ? orderUser.U_CREATED_ON.ToString() : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("SAMPLE");
                            writer.WriteElementString("SAMPLE_ID",sample.SAMPLE_ID.ToString());
                            writer.WriteElementString("PRIORITY",sample.PRIORITY.ToString());
                            writer.WriteElementString("ARCHIVED_CHILD_COMPLETE",sample.ARCHIVED_CHILD_COMPLETE);
                            writer.WriteElementString("LOCATION_ID",sample.LOCATION_ID.ToString());
                            writer.WriteElementString("PRODUCT_ID",sample.PRODUCT_ID.ToString());
                            writer.WriteElementString("GROUP_ID",sample.GROUP_ID.ToString());
                            writer.WriteElementString("SDG_ID",sample.SDG_ID.ToString());
                            writer.WriteElementString("PREVIOUS_SAMPLE",sample.PREVIOUS_SAMPLE.ToString());
                            writer.WriteElementString("OPERATOR_ID",sample.OPERATOR_ID.ToString());
                            writer.WriteElementString("CLIENT_ID",sample.CLIENT_ID.ToString());
                            writer.WriteElementString("SAMPLED_ON",sample.SAMPLED_ON.ToString());
                            writer.WriteElementString("SAMPLED_BY",sample.SAMPLED_BY.ToString());
                            writer.WriteElementString("DATE_RESULTS_REQUIRED",sample.DATE_RESULTS_REQUIRED.ToString());
                            writer.WriteElementString("EXTERNAL_REFERENCE",sample.EXTERNAL_REFERENCE);
                            writer.WriteElementString("CONCLUSION",sample.CONCLUSION);
                            writer.WriteElementString("EXPECTED_ON",sample.EXPECTED_ON.ToString());
                            writer.WriteElementString("SAMPLE_TYPE",sample.SAMPLE_TYPE);
                            writer.WriteElementString("INSPECTION_PLAN_ID",sample.INSPECTION_PLAN_ID.ToString());
                            writer.WriteElementString("SAMPLE_TEMPLATE_ID",sample.SAMPLE_TEMPLATE_ID.ToString());
                            writer.WriteElementString("WORKFLOW_NODE_ID",sample.WORKFLOW_NODE_ID.ToString());
                            writer.WriteElementString("NAME",sample.NAME);
                            writer.WriteElementString("DESCRIPTION",sample.DESCRIPTION);
                            writer.WriteElementString("STATUS",sample.STATUS);
                            writer.WriteElementString("OLD_STATUS",sample.OLD_STATUS);
                            writer.WriteElementString("CREATED_ON",sample.CREATED_ON.ToString());
                            writer.WriteElementString("COMPLETED_ON",sample.COMPLETED_ON.ToString());
                            writer.WriteElementString("AUTHORISED_ON",sample.AUTHORISED_ON.ToString());
                            writer.WriteElementString("EVENTS",sample.EVENTS);
                            writer.WriteElementString("BLIND_SAMPLE",sample.BLIND_SAMPLE);
                            writer.WriteElementString("NEEDS_REVIEW",sample.NEEDS_REVIEW);
                            writer.WriteElementString("RECEIVED_BY",sample.RECEIVED_BY.ToString());
                            writer.WriteElementString("RECEIVED_ON",sample.RECEIVED_ON.ToString());
                            writer.WriteElementString("REPORTED",sample.REPORTED);
                            writer.WriteElementString("HAS_NOTES",sample.HAS_NOTES);
                            writer.WriteElementString("HAS_AUDITS",sample.HAS_AUDITS);
                            writer.WriteElementString("STUDY_ID",sample.STUDY_ID.ToString());
                            writer.WriteElementString("COFA_TEMPLATE_ID",sample.COFA_TEMPLATE_ID.ToString());
                            writer.WriteElementString("REVIEW_TEMPLATE_ID",sample.REVIEW_TEMPLATE_ID.ToString());
                            writer.WriteElementString("CREATED_BY",sample.CREATED_BY.ToString());
                            writer.WriteElementString("COMPLETED_BY",sample.COMPLETED_BY.ToString());
                            writer.WriteElementString("AUTHORISED_BY",sample.AUTHORISED_BY.ToString());
                            writer.WriteEndElement();

                            writer.WriteStartElement("SAMPLE_USER");
                            writer.WriteElementString("SAMPLE_ID",sampleUser.SAMPLE_ID.ToString());
                            writer.WriteElementString("U_ASSISTANT_MACRO",sampleUser.U_ASSISTANT_MACRO);
                            writer.WriteElementString("U_DATE_ARRIVAL",sampleUser.U_DATE_ARRIVAL.ToString());
                            writer.WriteElementString("U_DECAL",sampleUser.U_DECAL);
                            writer.WriteElementString("U_LIQUID_TYPE",sampleUser.U_LIQUID_TYPE);
                            writer.WriteElementString("U_MARK",sampleUser.U_MARK);
                            writer.WriteElementString("U_MARK_AS",sampleUser.U_MARK_AS);
                            writer.WriteElementString("U_MATERIAL",sampleUser.U_MATERIAL);
                            writer.WriteElementString("U_PATHOLOG_MACRO",sampleUser.U_PATHOLOG_MACRO);
                            writer.WriteElementString("U_MICRO_HEADER",sampleUser.U_MICRO_HEADER);
                            writer.WriteElementString("U_ORGAN",sampleUser.U_ORGAN);
                            writer.WriteElementString("U_ORGAN_CODE",sampleUser.U_ORGAN_CODE);
                            writer.WriteElementString("U_ORGAN_ID",sampleUser.U_ORGAN_ID.ToString());
                            writer.WriteElementString("U_ORGAN_SIDE",sampleUser.U_ORGAN_SIDE);
                            writer.WriteElementString("U_ORGAN_TOPOGRAPHY",sampleUser.U_ORGAN_TOPOGRAPHY);
                            writer.WriteElementString("U_OVERNIGHT",sampleUser.U_OVERNIGHT);
                            writer.WriteElementString("U_SAMPLE_CODE",sampleUser.U_SAMPLE_CODE);
                            writer.WriteElementString("U_SAMPLE_FIX",sampleUser.U_SAMPLE_FIX);
                            writer.WriteElementString("U_SAMPLE_TYPE",sampleUser.U_SAMPLE_TYPE);
                            writer.WriteElementString("U_VOLUME",sampleUser.U_VOLUME.ToString());
                            writer.WriteElementString("U_PROCEDURE_CODE",sampleUser.U_PROCEDURE_CODE);
                            writer.WriteElementString("U_PROCEDURE_ID",sampleUser.U_PROCEDURE_ID.ToString());
                            writer.WriteElementString("U_ORGAN_REMARK",sampleUser.U_ORGAN_REMARK);
                            writer.WriteElementString("U_TOPOGRAPHY",sampleUser.U_TOPOGRAPHY);
                            writer.WriteElementString("U_TOPOGRAPHY_CODE",sampleUser.U_TOPOGRAPHY_CODE);
                            writer.WriteElementString("U_ARCHIVE",sampleUser.U_ARCHIVE);
                            writer.WriteElementString("U_ORDER",sampleUser.U_ORDER.ToString());
                            writer.WriteElementString("U_HISTOLOGY_SIZE",sampleUser.U_HISTOLOGY_SIZE);
                            writer.WriteElementString("U_TEST_CODE",sampleUser.U_TEST_CODE);
                            writer.WriteElementString("U_CYTOLOGY_NEXT_STEP",sampleUser.U_CYTOLOGY_NEXT_STEP);
                            writer.WriteElementString("U_CYTOLOGY_SLIDE_TYPE",sampleUser.U_CYTOLOGY_SLIDE_TYPE);
                            writer.WriteElementString("U_COLOR",sampleUser.U_COLOR);
                            writer.WriteElementString("U_PATHOLAB_SAMPLE_NAME",sampleUser.U_PATHOLAB_SAMPLE_NAME);
                            writer.WriteElementString("U_TUMOR_SIZE",sampleUser.U_TUMOR_SIZE.ToString());
                            writer.WriteElementString("U_ASSUTA_NUMBER",sampleUser.U_ASSUTA_NUMBER);
                            writer.WriteEndElement();

                            writer.WriteStartElement("ALIQUOT");
                            writer.WriteElementString("ALIQUOT_ID",aliquot.ALIQUOT_ID.ToString());
                            writer.WriteElementString("SAMPLE_ID",aliquot.SAMPLE_ID.ToString());
                            writer.WriteElementString("OPERATOR_ID",aliquot.OPERATOR_ID.ToString());
                            writer.WriteElementString("LOCATION_ID",aliquot.LOCATION_ID.ToString());
                            writer.WriteElementString("GROUP_ID",aliquot.GROUP_ID.ToString());
                            writer.WriteElementString("PRIORITY",aliquot.PRIORITY.ToString());
                            writer.WriteElementString("MATRIX_TYPE",aliquot.MATRIX_TYPE);
                            writer.WriteElementString("CONCLUSION",aliquot.CONCLUSION);
                            writer.WriteElementString("CONDITION",aliquot.CONDITION);
                            writer.WriteElementString("AMOUNT",aliquot.AMOUNT.ToString());
                            writer.WriteElementString("DATE_RESULTS_REQUIRED",aliquot.DATE_RESULTS_REQUIRED.ToString());
                            writer.WriteElementString("EXPECTED_ON",aliquot.EXPECTED_ON.ToString());
                            writer.WriteElementString("EXPIRES_ON",aliquot.EXPIRES_ON.ToString());
                            writer.WriteElementString("PLATE_ALIQUOT_TYPE",aliquot.PLATE_ALIQUOT_TYPE.ToString());
                            writer.WriteElementString("ARCHIVED_CHILD_COMPLETE",aliquot.ARCHIVED_CHILD_COMPLETE);
                            writer.WriteElementString("CONTAINER_TYPE_ID",aliquot.CONTAINER_TYPE_ID.ToString());
                            writer.WriteElementString("LIQUOT_TEMPLATE_ID",aliquot.ALIQUOT_TEMPLATE_ID.ToString());
                            writer.WriteElementString("WORKFLOW_NODE_ID",aliquot.WORKFLOW_NODE_ID.ToString());
                            writer.WriteElementString("USAGE_COUNT",aliquot.USAGE_COUNT.ToString());
                            writer.WriteElementString("STOCK_TEMPLATE_ID",aliquot.STOCK_TEMPLATE_ID.ToString());
                            writer.WriteElementString("NAME",aliquot.NAME);
                            writer.WriteElementString("DESCRIPTION",aliquot.DESCRIPTION);
                            writer.WriteElementString("STATUS",aliquot.STATUS);
                            writer.WriteElementString("OLD_STATUS",aliquot.OLD_STATUS);
                            writer.WriteElementString("CREATED_ON",aliquot.CREATED_ON.ToString());
                            writer.WriteElementString("COMPLETED_ON",aliquot.COMPLETED_ON.ToString());
                            writer.WriteElementString("AUTHORISED_ON",aliquot.AUTHORISED_ON.ToString());
                            writer.WriteElementString("EVENTS",aliquot.EVENTS);
                            writer.WriteElementString("UNIT_ID",aliquot.UNIT_ID.ToString());
                            writer.WriteElementString("NEEDS_REVIEW",aliquot.NEEDS_REVIEW);
                            writer.WriteElementString("INSPECTION_PLAN_ID",aliquot.INSPECTION_PLAN_ID.ToString());
                            writer.WriteElementString("RECEIVED_BY",aliquot.RECEIVED_BY.ToString());
                            writer.WriteElementString("RECEIVED_ON",aliquot.RECEIVED_ON.ToString());
                            writer.WriteElementString("REPORTED",aliquot.REPORTED);
                            writer.WriteElementString("HAS_NOTES",aliquot.HAS_NOTES);
                            writer.WriteElementString("HAS_AUDITS",aliquot.HAS_AUDITS);
                            writer.WriteElementString("ALIQUOT_TYPE",aliquot.ALIQUOT_TYPE);
                            writer.WriteElementString("STORAGE",aliquot.STORAGE);
                            writer.WriteElementString("EXTERNAL_REFERENCE",aliquot.EXTERNAL_REFERENCE);
                            writer.WriteElementString("CREATED_BY",aliquot.CREATED_BY.ToString());
                            writer.WriteElementString("COMPLETED_BY",aliquot.COMPLETED_BY.ToString());
                            writer.WriteElementString("AUTHORISED_BY",aliquot.AUTHORISED_BY.ToString());
                            writer.WriteElementString("SUPPLIER_ID",aliquot.SUPPLIER_ID.ToString());
                            writer.WriteElementString("CHEMICAL_ID",aliquot.CHEMICAL_ID.ToString());
                            writer.WriteElementString("STOCK_TYPE_ID",aliquot.STOCK_TYPE_ID.ToString());
                            writer.WriteElementString("GRADE",aliquot.GRADE);
                            writer.WriteElementString("BATCH_NUMBER",aliquot.BATCH_NUMBER);
                            writer.WriteElementString("PURITY",aliquot.PURITY.ToString());
                            writer.WriteElementString("PLATE_ID",aliquot.PLATE_ID.ToString());
                            writer.WriteElementString("PLATE_ORDER",aliquot.PLATE_ORDER.ToString());
                            writer.WriteElementString("PLATE_ROW",aliquot.PLATE_ROW.ToString());
                            writer.WriteElementString("PLATE_COLUMN",aliquot.PLATE_COLUMN.ToString());
                            writer.WriteElementString("PLATE_EDITOR_ID",aliquot.PLATE_EDITOR_ID.ToString());
                            writer.WriteEndElement();

                            writer.WriteStartElement("ALIQUOT_USER");
                            writer.WriteElementString("ALIQUOT_ID",aliquotUser.ALIQUOT_ID.ToString());
                            writer.WriteElementString("U_ALIQUOT_FIX",aliquotUser.U_ALIQUOT_FIX);
                            writer.WriteElementString("U_ALIQUOT_REMARK",aliquotUser.U_ALIQUOT_REMARK);
                            writer.WriteElementString("U_ALIQUOT_STATION",aliquotUser.U_ALIQUOT_STATION);
                            writer.WriteElementString("U_ARCHIVE",aliquotUser.U_ARCHIVE);
                            writer.WriteElementString("U_COLOR_TYPE",aliquotUser.U_COLOR_TYPE);
                            writer.WriteElementString("U_DECAL",aliquotUser.U_DECAL);
                            writer.WriteElementString("U_EXPRESS",aliquotUser.U_EXPRESS);
                            writer.WriteElementString("U_FORMATTED_COLOR",aliquotUser.U_FORMATTED_COLOR);
                            writer.WriteElementString("U_FROZEN",aliquotUser.U_FROZEN);
                            writer.WriteElementString("U_GLASS_TYPE",aliquotUser.U_GLASS_TYPE);
                            writer.WriteElementString("U_IS_CELL_BLOCK",aliquotUser.U_IS_CELL_BLOCK);
                            writer.WriteElementString("U_LAST_ENTRANCE",aliquotUser.U_LAST_ENTRANCE.ToString());
                            writer.WriteElementString("U_LAST_LABORANT",aliquotUser.U_LAST_LABORANT.ToString());
                            writer.WriteElementString("U_LAST_MICROTOM",aliquotUser.U_LAST_MICROTOM.ToString());
                            writer.WriteElementString("U_LOCATION",aliquotUser.U_LOCATION);
                            writer.WriteElementString("U_NUM_OF_TISSUES",aliquotUser.U_NUM_OF_TISSUES);
                            writer.WriteElementString("U_OLD_ALIQUOT_STATION",aliquotUser.U_OLD_ALIQUOT_STATION);
                            writer.WriteElementString("U_OVERNIGHT",aliquotUser.U_OVERNIGHT);
                            writer.WriteElementString("U_SETTING_UP",aliquotUser.U_SETTING_UP);
                            writer.WriteElementString("U_AUTOTEC",aliquotUser.U_AUTOTEC);
                            writer.WriteElementString("U_BLOCK_NAME",aliquotUser.U_BLOCK_NAME);
                            writer.WriteElementString("U_LAST_MICROTOME",aliquotUser.U_LAST_MICROTOME.ToString());
                            writer.WriteElementString("U_PRINTED_ON",aliquotUser.U_PRINTED_ON.ToString());
                            writer.WriteElementString("U_PRINTER_COL",aliquotUser.U_PRINTER_COL);
                            writer.WriteElementString("U_BASKET",aliquotUser.U_BASKET);
                            writer.WriteElementString("U_SEND_TO_PATHOLOG_ON",aliquotUser.U_SEND_TO_PATHOLOG_ON.ToString());
                            writer.WriteElementString("U_BACK_FROM_PATHOLOG_ON",aliquotUser.U_BACK_FROM_PATHOLOG_ON.ToString());
                            writer.WriteElementString("U_PATHOLAB_ALIQUOT_NAME",aliquotUser.U_PATHOLAB_ALIQUOT_NAME);
                            writer.WriteElementString("U_CANCELED_ON",aliquotUser.U_CANCELED_ON.ToString());
                            writer.WriteElementString("U_CALCULATE_DEBIT",aliquotUser.U_CALCULATE_DEBIT);
                            writer.WriteElementString("U_DEBIT_ID",aliquotUser.U_DEBIT_ID.ToString());
                            writer.WriteElementString("U_EXTERNAL_LAB_NUM",aliquotUser.U_EXTERNAL_LAB_NUM);
                            writer.WriteElementString("U_AGREE_TYPE",aliquotUser.U_AGREE_TYPE);
                            writer.WriteElementString("U_HISTOKINT_BATCH",aliquotUser.U_HISTOKINT_BATCH);
                            writer.WriteEndElement();

                            writer.WriteStartElement("TEST");
                            writer.WriteElementString("TEST_ID",test.TEST_ID.ToString());
                            writer.WriteElementString("ALIQUOT_ID",test.ALIQUOT_ID.ToString());
                            writer.WriteElementString("REPLICATE_NUMBER",test.REPLICATE_NUMBER.ToString());
                            writer.WriteElementString("GROUP_ID",test.GROUP_ID.ToString());
                            writer.WriteElementString("OPERATOR_ID",test.OPERATOR_ID.ToString());
                            writer.WriteElementString("LOCATION_ID",test.LOCATION_ID.ToString());
                            writer.WriteElementString("INSTRUMENT_ID",test.INSTRUMENT_ID.ToString());
                            writer.WriteElementString("PRIORITY",test.PRIORITY.ToString());
                            writer.WriteElementString("DATE_RESULTS_REQUIRED",test.DATE_RESULTS_REQUIRED.ToString());
                            writer.WriteElementString("EXPECTED_ON",test.EXPECTED_ON.ToString());
                            writer.WriteElementString("CONCLUSION",test.CONCLUSION);
                            writer.WriteElementString("WORKFLOW_NODE_ID",test.WORKFLOW_NODE_ID.ToString());
                            writer.WriteElementString("TEST_TEMPLATE_ID",test.TEST_TEMPLATE_ID.ToString());
                            writer.WriteElementString("NAME",test.NAME);
                            writer.WriteElementString("DESCRIPTION",test.DESCRIPTION);
                            writer.WriteElementString("STATUS",test.STATUS);
                            writer.WriteElementString("OLD_STATUS",test.OLD_STATUS);
                            writer.WriteElementString("CREATED_ON",test.CREATED_ON.ToString());
                            writer.WriteElementString("COMPLETED_ON",test.COMPLETED_ON.ToString());
                            writer.WriteElementString("AUTHORISED_ON",test.AUTHORISED_ON.ToString());
                            writer.WriteElementString("EVENTS",test.EVENTS);
                            writer.WriteElementString("NEEDS_REVIEW",test.NEEDS_REVIEW);
                            writer.WriteElementString("INSPECTION_PLAN_ID",test.INSPECTION_PLAN_ID.ToString());
                            writer.WriteElementString("HAS_NOTES",test.HAS_NOTES);
                            writer.WriteElementString("HAS_AUDITS",test.HAS_AUDITS);
                            writer.WriteElementString("STARTED_ON",test.STARTED_ON.ToString());
                            writer.WriteElementString("CREATED_BY",test.CREATED_BY.ToString());
                            writer.WriteElementString("COMPLETED_BY",test.COMPLETED_BY.ToString());
                            writer.WriteElementString("AUTHORISED_BY",test.AUTHORISED_BY.ToString());
                            writer.WriteElementString("PLATE_ID",test.PLATE_ID.ToString());
                            writer.WriteEndElement();

                            writer.WriteStartElement("RESULT");
                            writer.WriteElementString("RESULT_ID",result.RESULT_ID.ToString());
                            writer.WriteElementString("WORKSHEET_SESSION_ID",result.WORKSHEET_SESSION_ID.ToString());
                            writer.WriteElementString("RESULT_TYPE",result.RESULT_TYPE);
                            writer.WriteElementString("DILUTION_FACTOR",result.DILUTION_FACTOR.ToString());
                            writer.WriteElementString("FORMATTED_RESULT",result.FORMATTED_RESULT);
                            writer.WriteElementString("AQC_FAILURE",result.AQC_FAILURE);
                            writer.WriteElementString("INSTRUMENT_ID",result.INSTRUMENT_ID.ToString());
                            writer.WriteElementString("INSTRUMENT_FILE_ID",result.INSTRUMENT_FILE_ID.ToString());
                            writer.WriteElementString("ORIGINAL_RESULT",result.ORIGINAL_RESULT);
                            writer.WriteElementString("RAW_NUMERIC_RESULT",result.RAW_NUMERIC_RESULT.ToString());
                            writer.WriteElementString("RAW_DATETIME_RESULT",result.RAW_DATETIME_RESULT.ToString());
                            writer.WriteElementString("CALCULATED_NUMERIC_RESULT",result.CALCULATED_NUMERIC_RESULT.ToString());
                            writer.WriteElementString("RESULT_PREFIX",result.RESULT_PREFIX);
                            writer.WriteElementString("CONCLUSION",result.CONCLUSION);
                            writer.WriteElementString("AD_HOC",result.AD_HOC);
                            writer.WriteElementString("FORMATTED_UNIT",result.FORMATTED_UNIT);
                            writer.WriteElementString("ORIGINAL_UNIT",result.ORIGINAL_UNIT);
                            writer.WriteElementString("MODIFIED",result.MODIFIED);
                            writer.WriteElementString("OPERATOR_ID",result.OPERATOR_ID.ToString());
                            writer.WriteElementString("WORKFLOW_NODE_ID",result.WORKFLOW_NODE_ID.ToString());
                            writer.WriteElementString("RESULT_TEMPLATE_ID",result.RESULT_TEMPLATE_ID.ToString());
                            writer.WriteElementString("TEST_ID",result.TEST_ID.ToString());
                            writer.WriteElementString("WORKSHEET_ID",result.WORKSHEET_ID.ToString());
                            writer.WriteElementString("NAME",result.NAME);
                            writer.WriteElementString("DESCRIPTION",result.DESCRIPTION);
                            writer.WriteElementString("STATUS",result.STATUS);
                            writer.WriteElementString("OLD_STATUS",result.OLD_STATUS);
                            writer.WriteElementString("CREATED_ON",result.CREATED_ON.ToString());
                            writer.WriteElementString("COMPLETED_ON",result.COMPLETED_ON.ToString());
                            writer.WriteElementString("AUTHORISED_ON",result.AUTHORISED_ON.ToString());
                            writer.WriteElementString("EVENTS",result.EVENTS);
                            writer.WriteElementString("NEEDS_REVIEW",result.NEEDS_REVIEW);
                            writer.WriteElementString("INSPECTION_PLAN_ID",result.INSPECTION_PLAN_ID.ToString());
                            writer.WriteElementString("HAS_NOTES",result.HAS_NOTES);
                            writer.WriteElementString("HAS_AUDITS",result.HAS_AUDITS);
                            writer.WriteElementString("REPORTED",result.REPORTED);
                            writer.WriteElementString("ANALYTE_ID",result.ANALYTE_ID.ToString());
                            writer.WriteElementString("CREATED_BY",result.CREATED_BY.ToString());
                            writer.WriteElementString("COMPLETED_BY",result.COMPLETED_BY.ToString());
                            writer.WriteElementString("AUTHORISED_BY",result.AUTHORISED_BY.ToString());
                            writer.WriteElementString("MAXIMUM",result.MAXIMUM.ToString());
                            writer.WriteElementString("MINIMUM",result.MINIMUM.ToString());
                            writer.WriteElementString("ROUNDING_METHOD",result.ROUNDING_METHOD);
                            writer.WriteEndElement();


                            //טבלה לא קיימת
                            writer.WriteStartElement("RESULT_USER");
                            writer.WriteElementString("RESULT_ID","");
                            writer.WriteElementString("U_VISIBLE","");
                            writer.WriteElementString("U_BOLD","");
                            writer.WriteElementString("U_ORDER","");
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_CUSTOMER");
                            writer.WriteElementString("U_CUSTOMER_ID", customer != null ? customer.U_CUSTOMER_ID.ToString() : "");
                            writer.WriteElementString("NAME", customer != null ? customer.NAME : "");
                            writer.WriteElementString("DESCRIPTION", customer != null ? customer.DESCRIPTION : "");
                            writer.WriteElementString("VERSION", customer != null ? customer.VERSION : "");
                            writer.WriteElementString("VERSION_STATUS", customer != null ? customer.VERSION_STATUS : "");
                            writer.WriteElementString("GROUP_ID", customer!= null ? customer.GROUP_ID.ToString() : "");
                            writer.WriteElementString("PARENT_VERSION_ID", customer != null ? customer.PARENT_VERSION_ID.ToString() : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_CUSTOMER_USER");
                            writer.WriteElementString("U_CUSTOMER_ID", customerUser != null ? customerUser.U_CUSTOMER_ID.ToString() : "");
                            writer.WriteElementString("U_CUSTOMER_CODE", customerUser != null ? customerUser.U_CUSTOMER_CODE : "");
                            writer.WriteElementString("U_PRICE_LIST_NAME", customerUser != null ? customerUser.U_PRICE_LIST_NAME : "");
                            writer.WriteElementString("U_CUSTOMER_LTN_NAME", customerUser != null ? customerUser.U_CUSTOMER_LTN_NAME : "");
                            writer.WriteElementString("U_CUSTOMER_TYPE", customerUser != null ? customerUser.U_CUSTOMER_TYPE : "");
                            writer.WriteElementString("U_CUSTOMER_GROUP", customerUser != null ? customerUser.U_CUSTOMER_GROUP : "");
                            writer.WriteElementString("U_LETTER_B", customerUser != null ? customerUser.U_LETTER_B : "");
                            writer.WriteElementString("U_LETTER_C", customerUser != null ? customerUser.U_LETTER_C : "");
                            writer.WriteElementString("U_LETTER_P", customerUser != null ? customerUser.U_LETTER_P : "");
                            writer.WriteElementString("U_INC_VAT", customerUser != null ? customerUser.U_INC_VAT : "");
                            writer.WriteElementString("U_PAY_TRANSPORT", customerUser != null ? customerUser.U_PAY_TRANSPORT : "");
                            writer.WriteElementString("U_SMALL_UP_TO", customerUser != null ? customerUser.U_SMALL_UP_TO.ToString() : "");
                            writer.WriteElementString("U_MEDIUM_UP_TO", customerUser != null ? customerUser.U_MEDIUM_UP_TO.ToString() : "");
                            writer.WriteElementString("U_LARGE_UP_TO", customerUser != null ? customerUser.U_LARGE_UP_TO.ToString() : "");
                            writer.WriteElementString("U_PRIVATE_LIBRARY", customerUser != null ? customerUser.U_PRIVATE_LIBRARY : "");
                            writer.WriteElementString("U_PRICE_LIST_ID", customerUser != null ? customerUser.U_PRICE_LIST_ID.ToString() : "");
                            writer.WriteElementString("U_PRICE_LIST_NAME", customerUser != null ? customerUser.U_PRICE_LIST_NAME : "");
                            writer.WriteElementString("U_DISTRICT", customerUser != null ? customerUser.U_DISTRICT : "");
                            writer.WriteElementString("U_GRP_CLINIC_CODE", customerUser != null ? customerUser.U_GRP_CLINIC_CODE : "");
                            writer.WriteElementString("U_FAX_NBR", customerUser != null ? customerUser.U_FAX_NBR : "");
                            writer.WriteElementString("U_EMAIL_ADDRESS", customerUser != null ? customerUser.U_EMAIL_ADDRESS : "");
                            writer.WriteElementString("U_CLALIT", customerUser != null ? customerUser.U_CLALIT : "");
                            writer.WriteElementString("U_ASCII_TYPE", customerUser != null ? customerUser.U_ASCII_TYPE : "");
                            writer.WriteElementString("U_PRICE_LEVEL", customerUser != null ? customerUser.U_PRICE_LEVEL : "");
                            writer.WriteElementString("U_SEND_CODED_LETTER", customerUser != null ? customerUser.U_SEND_CODED_LETTER : "");
                            writer.WriteElementString("U_XL_UP_TO", customerUser != null ? customerUser.U_XL_UP_TO.ToString() : "");
                            writer.WriteElementString("U_SEND_MAIL_OPTION","");//שדה לא קיים
                            writer.WriteElementString("U_XXL_UP_TO","");//שדה לא קיים
                            writer.WriteEndElement();

                            writer.WriteStartElement("SUPPLIER");
                            writer.WriteElementString("SUPPLIER_ID", supplier != null ? supplier.SUPPLIER_ID.ToString() : "");
                            writer.WriteElementString("GROUP_ID", supplier != null ? supplier.GROUP_ID.ToString() : "");
                            writer.WriteElementString("NAME",supplier != null ? supplier.NAME : "");
                            writer.WriteElementString("VERSION", supplier != null ? supplier.VERSION : "");
                            writer.WriteElementString("DESCRIPTION", supplier != null ? supplier.DESCRIPTION : "");
                            writer.WriteElementString("VERSION_STATUS", supplier != null ? supplier.VERSION_STATUS : "");
                            writer.WriteElementString("PARENT_VERSION_ID", supplier != null ? supplier.PARENT_VERSION_ID.ToString() : "");
                            writer.WriteElementString("SUPPLIER_TYPE", supplier != null ? supplier.SUPPLIER_TYPE : "");
                            writer.WriteElementString("SUPPLIER_CODE", supplier != null ? supplier.SUPPLIER_CODE : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("SUPPLIER_USER");
                            writer.WriteElementString("SUPPLIER_ID", supplierUser != null ? supplierUser.SUPPLIER_ID.ToString() : "");
                            writer.WriteElementString("U_DEGREE",supplierUser != null ? supplierUser.U_DEGREE : "");
                            writer.WriteElementString("U_FIRST_NAME", supplierUser != null ? supplierUser.U_FIRST_NAME : "");
                            writer.WriteElementString("U_LAST_NAME", supplierUser != null ? supplierUser.U_LAST_NAME : "");
                            writer.WriteElementString("U_ID_NBR", supplierUser != null ? supplierUser.U_ID_NBR : "");
                            writer.WriteElementString("U_LICENSE_NBR", supplierUser != null ? supplierUser.U_LICENSE_NBR : "");
                            writer.WriteElementString("U_PROFICENCY", supplierUser != null ? supplierUser.U_PROFICENCY : "");
                            writer.WriteElementString("U_EMAIL_ADDRESS", supplierUser != null ? supplierUser.U_EMAIL_ADDRESS : "");
                            writer.WriteElementString("U_SEND_CODED_LETTER", supplierUser != null ? supplierUser.U_SEND_CODED_LETTER : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_CONTAINER");
                            writer.WriteElementString("U_CONTAINER_ID", container != null ? container.U_CONTAINER_ID.ToString() : "");
                            writer.WriteElementString("NAME",container != null ? container.NAME : "");
                            writer.WriteElementString("DESCRIPTION", container != null ? container.DESCRIPTION : "");
                            writer.WriteElementString("VERSION", container != null ? container.VERSION : "");
                            writer.WriteElementString("VERSION_STATUS", container != null ? container.VERSION_STATUS : "");
                            writer.WriteElementString("GROUP_ID", container != null ? container.GROUP_ID.ToString() : "");
                            writer.WriteElementString("PARENT_VERSION_ID", container != null ? container.PARENT_VERSION_ID.ToString() : "");
                            writer.WriteElementString("TEMPLATE_ID", container != null ? container.TEMPLATE_ID.ToString() : "");
                            writer.WriteElementString("WORKFLOW_NODE_ID", container != null ? container.WORKFLOW_NODE_ID.ToString() : "");
                            writer.WriteElementString("EVENTS",container != null ? container.EVENTS : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_CONTAINER_USER");
                            writer.WriteElementString("U_CONTAINER_ID", containerUser != null ? containerUser.U_CONTAINER_ID.ToString() : "");
                            writer.WriteElementString("U_RECEIVE_NUMBER",containerUser != null ? containerUser.U_RECEIVE_NUMBER : "");
                            writer.WriteElementString("U_RECEIVED_ON", containerUser != null ? containerUser.U_RECEIVED_ON.ToString() : "");
                            writer.WriteElementString("U_SEND_ON", containerUser != null ? containerUser.U_SEND_ON.ToString() : "");
                            writer.WriteElementString("U_CLINIC", containerUser != null ? containerUser.U_CLINIC.ToString() : "");
                            writer.WriteElementString("U_NUMBER_OF_ORDERS", containerUser != null ? containerUser.U_NUMBER_OF_ORDERS.ToString() : "");
                            writer.WriteElementString("U_NUMBER_OF_SAMPLES", containerUser != null ? containerUser.U_NUMBER_OF_SAMPLES.ToString() : "");
                            writer.WriteElementString("U_CREATE_BY", containerUser != null ? containerUser.U_CREATE_BY.ToString() : "");
                            writer.WriteElementString("U_STATUS", containerUser != null ? containerUser.U_STATUS : "");
                            writer.WriteElementString("U_FAX_SEND_ON", containerUser != null ? containerUser.U_FAX_SEND_ON.ToString() : "");
                            writer.WriteElementString("U_CONTEINER_ID", containerUser != null ? containerUser.U_CONTEINER_ID.ToString() : "");
                            writer.WriteElementString("U_BIG_QNT", containerUser != null ? containerUser.U_BIG_QNT.ToString() : "");
                            writer.WriteElementString("U_URGENT_QNT", containerUser != null ? containerUser.U_URGENT_QNT.ToString() : "");
                            writer.WriteElementString("U_ALL_ARRIVED","");//שדה לא קיים
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_DEBIT");
                            writer.WriteElementString("U_DEBIT_ID", debit != null ? debit.U_DEBIT_ID.ToString() : "");
                            writer.WriteElementString("NAME",debit != null ? debit.NAME : "");
                            writer.WriteElementString("DESCRIPTION", debit != null ? debit.DESCRIPTION : "");
                            writer.WriteElementString("VERSION", debit != null ? debit.VERSION : "");
                            writer.WriteElementString("VERSION_STATUS", debit != null ? debit.VERSION_STATUS : "");
                            writer.WriteElementString("GROUP_ID", debit != null ? debit.GROUP_ID.ToString() : "");
                            writer.WriteElementString("PARENT_VERSION_ID", debit != null ? debit.PARENT_VERSION_ID.ToString() : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_DEBIT_USER");
                            writer.WriteElementString("U_DEBIT_ID", debitUser != null ? debitUser.U_DEBIT_ID.ToString() : "");
                            writer.WriteElementString("U_ORDER_ID", debitUser != null ? debitUser.U_ORDER_ID.ToString() : "");
                            writer.WriteElementString("U_SDG_NAME",debitUser != null ? debitUser.U_SDG_NAME : "");
                            writer.WriteElementString("U_EVENT_DATE", debitUser != null ? debitUser.U_EVENT_DATE.ToString() : "");
                            writer.WriteElementString("U_PARTS_ID", debitUser != null ? debitUser.U_PARTS_ID.ToString() : "");
                            writer.WriteElementString("U_PART_TEXT", debitUser != null ? debitUser.U_PART_TEXT : "");
                            writer.WriteElementString("U_PART_PRICE", debitUser != null ? debitUser.U_PART_PRICE.ToString() : "");
                            writer.WriteElementString("U_PRICE_INC_VAT", debitUser != null ? debitUser.U_PRICE_INC_VAT.ToString() : "");
                            writer.WriteElementString("U_QUANTITY", debitUser != null ? debitUser.U_QUANTITY.ToString() : "");
                            writer.WriteElementString("U_LINE_AMOUNT", debitUser != null ? debitUser.U_LINE_AMOUNT.ToString() : "");
                            writer.WriteElementString("U_ENTITY_ID", debitUser != null ? debitUser.U_ENTITY_ID.ToString() : "");
                            writer.WriteElementString("U_DEBIT_STATUS", debitUser != null ? debitUser.U_DEBIT_STATUS : "");
                            writer.WriteElementString("U_INVC_NBR", debitUser != null ? debitUser.U_INVC_NBR : "");
                            writer.WriteElementString("U_INVC_DATE", debitUser != null ? debitUser.U_INVC_DATE.ToString() : "");
                            writer.WriteElementString("U_LAST_UPDATE", debitUser != null ? debitUser.U_LAST_UPDATE.ToString() : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_CLINIC");
                            writer.WriteElementString("U_CLINIC_ID", clinic != null ? clinic.U_CLINIC_ID.ToString() : "");
                            writer.WriteElementString("NAME",clinic != null ? clinic.NAME : "");
                            writer.WriteElementString("DESCRIPTION", clinic != null ? clinic.DESCRIPTION : "");
                            writer.WriteElementString("VERSION", clinic != null ? clinic.VERSION : "");
                            writer.WriteElementString("VERSION_STATUS", clinic != null ? clinic.VERSION_STATUS : "");
                            writer.WriteElementString("GROUP_ID", clinic != null ? clinic.GROUP_ID.ToString() : "");
                            writer.WriteElementString("PARENT_VERSION_ID", clinic != null ? clinic.PARENT_VERSION_ID.ToString() : "");
                            writer.WriteEndElement();

                            writer.WriteStartElement("U_CLINIC_USER");
                            writer.WriteElementString("U_CLINIC_ID", clinicUser != null ? clinicUser.U_CLINIC_ID.ToString() : "");
                            writer.WriteElementString("U_DISTRICT",clinicUser != null ? clinicUser.U_DISTRICT : "");
                            writer.WriteElementString("U_FAX_NBR", clinicUser != null ? clinicUser.U_FAX_NBR : "");
                            writer.WriteElementString("U_CLINIC_NAME", clinicUser != null ? clinicUser.U_CLINIC_NAME : "");
                            writer.WriteElementString("U_EX_NUMBER", clinicUser != null ? clinicUser.U_EX_NUMBER : "");
                            writer.WriteElementString("U_GRP_CODE", clinicUser != null ? clinicUser.U_GRP_CODE : "");
                            writer.WriteElementString("U_GRP_NAME", clinicUser != null ? clinicUser.U_GRP_NAME : "");
                            writer.WriteElementString("U_CUSTOMER_CODE", clinicUser != null ? clinicUser.U_CUSTOMER_CODE : "");
                            writer.WriteElementString("U_CUSTOMER_ID", clinicUser != null ? clinicUser.U_CUSTOMER_ID.ToString() : "");
                            writer.WriteElementString("U_CLINIC_CODE", clinicUser != null ? clinicUser.U_CLINIC_CODE : "");
                            writer.WriteElementString("U_EMAIL_ADDRESS", clinicUser != null ? clinicUser.U_EMAIL_ADDRESS : "");
                            writer.WriteElementString("U_SENDER", clinicUser != null ? clinicUser.U_SENDER : "");
                            writer.WriteElementString("U_PRIVATE_LIBRARY", clinicUser != null ? clinicUser.U_PRIVATE_LIBRARY : "");
                            writer.WriteElementString("U_ASSUTA_CLINIC_CODE", clinicUser != null ? clinicUser.U_ASSUTA_CLINIC_CODE : "");
                            writer.WriteElementString("U_ASSUTA_DIVISION_CODE", clinicUser != null ? clinicUser.U_ASSUTA_DIVISION_CODE : "");
                            writer.WriteElementString("U_SEND_CODED_LETTER", clinicUser != null ? clinicUser.U_SEND_CODED_LETTER : "");
                            writer.WriteElementString("U_SEND_MAIL_OPTION","");//שדה לא קיים
                            writer.WriteEndElement();

                            writer.WriteEndElement();

                            writer.WriteEndDocument();
                            writer.Flush();
                            writer.Close();
                        }
                    }else{
                      // MessageBox.Show("is not slide");

                    }
                    

                }
                else if (tableName == "SDG")
                {
                }

            }
            catch(Exception e)
            {
                string err = e.Message;
            }
        }

        /// <summary>
        /// Init nautilus con
        /// </summary>
        /// <param name="ntlsCon"></param>
        /// <returns></returns>
        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {

            OracleConnection connection = null;

            if (ntlsCon != null)
            {


                // Initialize variables
                String roleCommand;
                // Try/Catch block
                try
                {
                    _connectionString = ntlsCon.GetADOConnectionString();

                    var splited = _connectionString.Split(';');

                    var cs = "";

                    for (int i = 1; i < splited.Count(); i++)
                    {
                        cs += splited[i] + ';';
                    }

                    var username = ntlsCon.GetUsername();
                    if (string.IsNullOrEmpty(username))
                    {
                        var serverDetails = ntlsCon.GetServerDetails();
                        cs = "User Id=/;Data Source=" + serverDetails + ";";
                    }

                    //Create the connection
                    connection = new OracleConnection(cs);

                    // Open the connection
                    connection.Open();

                    // Get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    // Set role lims user
                    if (limsUserPassword == "")
                    {
                        // LIMS_USER is not password protected
                        roleCommand = "set role lims_user";
                    }
                    else
                    {
                        // LIMS_USER is password protected.
                        roleCommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    // set the Oracle user for this connecition
                    OracleCommand command = new OracleCommand(roleCommand, connection);

                    // Try/Catch block
                    try
                    {
                        // Execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        // Throw the exception
                        throw new Exception("Inconsistent role Security : " + f.Message);
                    }

                    // Get the session id
                    sessionId = ntlsCon.GetSessionId();

                    // Connect to the same session
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    // Build the command
                    command = new OracleCommand(sSql, connection);

                    // Execute the command
                    command.ExecuteNonQuery();

                }
                catch (Exception e)
                {
                    // Throw the exception
                    throw e;
                }
                // Return the connection
            }
            return connection;
        }
    }
}
