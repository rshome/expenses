using PaystreamExpenses;

namespace PaystreamExpenses
{
    public class LatestWeek
    {
        public static void Main(string[] args)
        {
            Driver wd = new Driver();            

            wd.Login();
            wd.SelectAccountandExpenses();
            
            wd.Broadband();
            wd.DeclareExpensesDriving();

            //wd.DeclarePhoneCalls();

            wd.DeclareExpensesCoffee();
            wd.DeclareExpensesParking();
            wd.DeclareExpensesLunch();
            wd.DeclareExpensesBreakfast();
            //wd.WeekTrainPass();




        }
    }
}
