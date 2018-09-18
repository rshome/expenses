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
            wd.DeclarePhoneCalls();

            wd.DeclareExpensesCoffee();
            wd.DeclareExpensesParking();
            wd.DeclareExpensesLunch();
            wd.DeclareExpensesBreakfast();
            wd.WeekTrainPass();

            for (int i = 0; i < 5; i++)
            {
                wd.DeclareExpensesDriving();
            }

        }
    }
}
