using UiPath.CodedWorkflows;
using System;

namespace ExcelPractice
{
    public class ConnectionsManager
    {
        public DriveFactory Drive { get; set; }

        public GmailFactory Gmail { get; set; }

        public GoogleSheetsFactory GoogleSheets { get; set; }

        public ConnectionsManager(ICodedWorkflowsServiceContainer resolver)
        {
            Drive = new DriveFactory(resolver);
            Gmail = new GmailFactory(resolver);
            GoogleSheets = new GoogleSheetsFactory(resolver);
        }
    }
}