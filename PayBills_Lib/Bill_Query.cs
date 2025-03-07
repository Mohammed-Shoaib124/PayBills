using System.Data;
using QBFC16Lib;

namespace PayBills_Lib
{
    public class Bill_Query
    {

        public static List<Bill> DoBillToPayQuery()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;
            List<Bill> bills = new List<Bill>();

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildBillToPayQueryRq(requestMsgSet);

                //IBillToPayQuery BillToPayQueryRq = requestMsgSet.AppendBillToPayQueryRq();

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                Console.WriteLine($"Raw Response: {responseMsgSet.ToXMLString()}");

                //End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                bills = WalkBillToPayQueryRs(responseMsgSet);

                return bills;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
                Console.WriteLine("Stack Trace: " + e.StackTrace);
                if (sessionBegun)
                {
                    sessionManager.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager.CloseConnection();
                }
                return bills;
            }
        }
        public static void BuildBillToPayQueryRq(IMsgSetRequest requestMsgSet)
        {
            IBillToPayQuery BillToPayQueryRq = requestMsgSet.AppendBillToPayQueryRq();
            //Set attributes
            //Set field value for metaData
            BillToPayQueryRq.metaData.SetValue(ENmetaData.mdMetaDataAndResponseData);
            //Set field value for ListID
            BillToPayQueryRq.PayeeEntityRef.ListID.SetValue("");
            //Set field value for FullName
            BillToPayQueryRq.PayeeEntityRef.FullName.SetValue("");
            //Set field value for ListID
            BillToPayQueryRq.APAccountRef.ListID.SetValue("");
            //Set field value for FullName
            //Set field value for DueDate
            BillToPayQueryRq.DueDate.SetValue(DateTime.Parse("12/15/2007"));
            string ORCurrencyFilterElementType2045 = "ListIDList";
            if (ORCurrencyFilterElementType2045 == "ListIDList")
            {
                //Set field value for ListIDList
                //May create more than one of these if needed
                BillToPayQueryRq.CurrencyFilter.ORCurrencyFilter.ListIDList.Add("200000-1011023419");
            }
            if (ORCurrencyFilterElementType2045 == "FullNameList")
            {
                //Set field value for FullNameList
                //May create more than one of these if needed
                BillToPayQueryRq.CurrencyFilter.ORCurrencyFilter.FullNameList.Add("ab");
            }
            //Set field value for IncludeRetElementList
            //May create more than one of these if needed
            BillToPayQueryRq.IncludeRetElementList.Add("ab");
        }

        public static List<Bill> WalkBillToPayQueryRs(IMsgSetResponse responseMsgSet)
        {
            List<Bill> bills = new List<Bill>();
            if (responseMsgSet == null) return bills;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return bills;

            Console.WriteLine("Checking response message...");
            Console.WriteLine($"Response Count: {responseMsgSet.ResponseList.Count}");

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                Console.WriteLine($"Response Code: {response.StatusCode}, Message: {response.StatusMessage}");

                if (response.StatusCode == 0 && response.Detail != null)
                {
                    ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                    if (responseType == ENResponseType.rtBillToPayQueryRs)
                    {
                        IBillToPayRetList BillToPayRet = (IBillToPayRetList)response.Detail;
                        bills = WalkBillToPayRet(BillToPayRet);
                    }
                }
            }
            return bills;
        }





        public static List<Bill> WalkBillToPayRet(IBillToPayRetList BillToPayRet)
        {
            List<Bill> bills = new List<Bill>();
            if (BillToPayRet == null) return bills;
            //Go through all the elements of IBillToPayRetList

            for (int i = 0; i < BillToPayRet.Count; i++) 
            {
                var bill = BillToPayRet.GetAt(i);
                if (bill.ORBillToPayRet != null)
                {
                    if (bill.ORBillToPayRet.BillToPay != null)
                    {
                        if (bill.ORBillToPayRet.BillToPay != null)
                        {
                            //Get value of TxnID
                            string TransactionId = (string)bill.ORBillToPayRet.BillToPay.TxnID.GetValue();
                            //Get value of TxnType
                            //ENTxnType TxnType2048 = (ENTxnType)bill.ORBillToPayRet.BillToPay.TxnType.GetValue();
                            //Get value of ListID
                            string ListId = "";
                            if (bill.ORBillToPayRet.BillToPay.APAccountRef.ListID != null)
                            {
                                 ListId = (string)bill.ORBillToPayRet.BillToPay.APAccountRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            string FullName = "";
                            if (bill.ORBillToPayRet.BillToPay.APAccountRef.FullName != null)
                            {
                                 FullName = (string)bill.ORBillToPayRet.BillToPay.APAccountRef.FullName.GetValue();
                            }
                            //Get value of TxnDate
                            DateTime TxnDate = (DateTime)bill.ORBillToPayRet.BillToPay.TxnDate.GetValue();
                            //Get value of RefNumber
                            string RefNumber = "";
                            if (bill.ORBillToPayRet.BillToPay.RefNumber != null)
                            {
                                 RefNumber = (string)bill.ORBillToPayRet.BillToPay.RefNumber.GetValue();
                            }
                            //Get value of DueDate
                            DateTime DueDate   = DateTime.Now;
                            if (bill.ORBillToPayRet.BillToPay.DueDate != null)
                            {
                                DueDate = (DateTime)bill.ORBillToPayRet.BillToPay.DueDate.GetValue();
                            }
                            //Get value of AmountDue
                            double AmountDue = (double)bill.ORBillToPayRet.BillToPay.AmountDue.GetValue();
                            //if (bill.ORBillToPayRet.BillToPay.CurrencyRef != null)
                            //{
                            //    //Get value of ListID
                            //    if (bill.ORBillToPayRet.BillToPay.CurrencyRef.ListID != null)
                            //    {
                            //        string ListID2055 = (string)bill.ORBillToPayRet.BillToPay.CurrencyRef.ListID.GetValue();
                            //    }
                            //    //Get value of FullName
                            //    if (bill.ORBillToPayRet.BillToPay.CurrencyRef.FullName != null)
                            //    {
                            //        string FullName2056 = (string)bill.ORBillToPayRet.BillToPay.CurrencyRef.FullName.GetValue();
                            //    }
                            //}

                            Console.WriteLine("Bill");
                            
                            Bill newbill = new Bill(TransactionId,ListId,FullName,TxnDate, RefNumber,DueDate,AmountDue);

                            bills.Add(newbill);
                        }
                    }
                    //if (bill.ORBillToPayRet.CreditToApply != null)
                    //{
                    //    if (bill.ORBillToPayRet.CreditToApply != null)
                    //    {
                    //        //Get value of TxnID
                    //        string TxnID2059 = (string)BillToPayRet.ORBillToPayRet.CreditToApply.TxnID.GetValue();
                    //        //Get value of TxnType
                    //        ENTxnType TxnType2060 = (ENTxnType)BillToPayRet.ORBillToPayRet.CreditToApply.TxnType.GetValue();
                    //        //Get value of ListID
                    //        if (BillToPayRet.ORBillToPayRet.CreditToApply.APAccountRef.ListID != null)
                    //        {
                    //            string ListID2061 = (string)BillToPayRet.ORBillToPayRet.CreditToApply.APAccountRef.ListID.GetValue();
                    //        }
                    //        //Get value of FullName
                    //        if (BillToPayRet.ORBillToPayRet.CreditToApply.APAccountRef.FullName != null)
                    //        {
                    //            string FullName2062 = (string)BillToPayRet.ORBillToPayRet.CreditToApply.APAccountRef.FullName.GetValue();
                    //        }
                    //        //Get value of TxnDate
                    //        DateTime TxnDate2063 = (DateTime)BillToPayRet.ORBillToPayRet.CreditToApply.TxnDate.GetValue();
                    //        //Get value of RefNumber
                    //        if (BillToPayRet.ORBillToPayRet.CreditToApply.RefNumber != null)
                    //        {
                    //            string RefNumber2064 = (string)BillToPayRet.ORBillToPayRet.CreditToApply.RefNumber.GetValue();
                    //        }
                    //        //Get value of CreditRemaining
                    //        double CreditRemaining2065 = (double)BillToPayRet.ORBillToPayRet.CreditToApply.CreditRemaining.GetValue();
                    //        if (BillToPayRet.ORBillToPayRet.CreditToApply.CurrencyRef != null)
                    //        {
                    //            //Get value of ListID
                    //            if (BillToPayRet.ORBillToPayRet.CreditToApply.CurrencyRef.ListID != null)
                    //            {
                    //                string ListID2066 = (string)BillToPayRet.ORBillToPayRet.CreditToApply.CurrencyRef.ListID.GetValue();
                    //            }
                    //            //Get value of FullName
                    //            if (BillToPayRet.ORBillToPayRet.CreditToApply.CurrencyRef.FullName != null)
                    //            {
                    //                string FullName2067 = (string)BillToPayRet.ORBillToPayRet.CreditToApply.CurrencyRef.FullName.GetValue();
                    //            }
                    //        }
                    //        //Get value of ExchangeRate
                    //        if (BillToPayRet.ORBillToPayRet.CreditToApply.ExchangeRate != null)
                    //        {
                    //            IQBFloatType ExchangeRate2068 = (IQBFloatType)BillToPayRet.ORBillToPayRet.CreditToApply.ExchangeRate.GetValue();
                    //        }
                    //        //Get value of CreditRemainingInHomeCurrency
                    //        if (BillToPayRet.ORBillToPayRet.CreditToApply.CreditRemainingInHomeCurrency != null)
                    //        {
                    //            double CreditRemainingInHomeCurrency2069 = (double)BillToPayRet.ORBillToPayRet.CreditToApply.CreditRemainingInHomeCurrency.GetValue();
                    //        }
                    //    }
                    //}
                }

            }


            return bills;
            
        }



    }
}
