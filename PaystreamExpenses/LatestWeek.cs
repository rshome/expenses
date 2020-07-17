using PaystreamExpenses;

namespace PaystreamExpenses
{
    public class LatestWeek : Driver
    {
        public static void Main(string[] args)
        {
            Driver wd = new Driver();            

            wd.Login();
            wd.SelectAccountandExpenses();

            wd.Broadband();
            wd.MonthlyTrainPass();
            //wd.Hotel();
            wd.DeclarePhoneCalls();

            wd.DeclareExpensesCoffee();
            wd.DeclareExpensesParking();

            wd.DeclareExpensesBreakfast();

            wd.DeclareExpensesLunch();
            wd.DeclareExpensesDriving();





        }

        
    }
}
