﻿using PaystreamExpenses;

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


            wd.MonthlyTrainPass();
            wd.DeclarePhoneCalls();

            wd.DeclareExpensesCoffee();
            wd.DeclareExpensesParking();

            wd.DeclareExpensesBreakfast();

            wd.DeclareExpensesDriving();

            wd.DeclareExpensesLunch();



        }

        
    }
}
