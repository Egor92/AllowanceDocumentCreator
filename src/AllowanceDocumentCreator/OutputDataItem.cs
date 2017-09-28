namespace AllowanceDocumentCreator
{
    public class OutputDataItem
    {
        #region Ctor

        public OutputDataItem(InputDataItem inputDataItem)
        {
            A = inputDataItem.A;
            B = inputDataItem.B;
            C = inputDataItem.C;
            D1 = inputDataItem.D1;
            D2 = inputDataItem.D2;
            E = inputDataItem.E;
        }

        #endregion

        #region LastName

        public string LastName
        {
            get { return "ЛЕВКОВИЧ"; }
        }

        #endregion

        #region FirstName

        public string FirstName
        {
            get { return "ОКСАНА"; }
        }

        #endregion

        #region FatherName

        public string FatherName
        {
            get { return "АЛЕКСАНДРОВНА"; }
        }

        #endregion

        #region DaysCount

        public int DaysCount
        {
            get { return 4; }
        }

        #endregion

        #region A

        public double A { get; set; }

        #endregion

        #region B

        public double B { get; set; }

        #endregion

        #region C

        public double C { get; set; }

        #endregion

        #region CPercent

        public double CPercent
        {
            get { return 22; }
        }

        #endregion
        
        #region D1

        public double D1 { get; set; }

        #endregion

        #region D1Percent

        public double D1Percent
        {
            get { return 2.9; }
        }

        #endregion

        #region D2

        public double D2 { get; set; }

        #endregion

        #region D2Percent

        public double D2Percent
        {
            get { return 0.2; }
        }

        #endregion

        #region E

        public double E { get; set; }

        #endregion

        #region EPercent

        public double EPercent
        {
            get { return 5.1; }
        }

        #endregion
    }
}
