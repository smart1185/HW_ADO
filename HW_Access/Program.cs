using System;
using ADOX;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.SqlClient;

namespace HW_Access
{
    class Program
    {
        //public static string Path = @"C:\Users\Lenovo\Documents\Visual Studio 2015\Projects\HW_Access\HW_Access\bin";
        //public static string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+Path+"; Jet OLEDB:Engine Type = 5; Integrated Security =SSPI;";
        public static string CONNECT_STRING = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=5",
               Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database.mdb"));


        //static void CreateOleDbCommand(string query1, string CONNECT_STRING)
        //{
        //    using (OleDbConnection connection = new OleDbConnection(CONNECT_STRING))
        //    {
        //        OleDbCommand cmd1 = new OleDbCommand(query1, connection);
        //        cmd1.Connection.Open();
        //        cmd1.ExecuteNonQuery();
        //        cmd1.Connection.Close();
        //    }
        //}
        static void Main(string[] args)
        {
            OleDbConnection connection = new OleDbConnection(CONNECT_STRING);
            OleDbCommand command1 = new OleDbCommand("CREATE TABLE TablesStopReason(intStopReason AUTOINCREMENT PRIMARY KEY, strReason string, bitDowntime bit, bitUnscheduled bit, bitPMStoppages bit, bitScheduledRepairsAndOther bit, intLocationId int NOT NULL)", connection);
            OleDbCommand command2 = new OleDbCommand("CREATE TABLE TablesManufacturer(intManufacturerID AUTOINCREMENT PRIMARY KEY, strManufactName string)", connection);
            OleDbCommand command3 = new OleDbCommand("CREATE TABLE TablesModel(intModelID AUTOINCREMENT PRIMARY KEY, strName string, intSMCSFamilyID int, strImage string, intManufacturerID int)", connection);

            if (File.Exists("Database.mdb"))
            {

                try
                {
                    connection.Open();
                    Console.WriteLine("Connection opened");
                    command1.ExecuteNonQuery();
                    command2.ExecuteNonQuery();
                    command3.ExecuteNonQuery();
                    
                }
                catch (Exception exc)
                {
                    Console.WriteLine(exc.Message);
                }
                finally
                {
                    connection.Close();
                    connection.Dispose();
                    Console.WriteLine("Connection closed");
                    if (command1 != null)
                        command1.Dispose();
                    if (command2 != null)
                        command2.Dispose();
                    if (command3 != null)
                        command3.Dispose();
                }
            }
            else
            {
                ADOX.Catalog catalog = new ADOX.Catalog();
                catalog.Create(CONNECT_STRING);
            }

            Console.WriteLine("Какую операцию вы хотите выполнить? 1 - внесение данных в таблицу, 2 - просмотр данных в таблицах, 3 - удаление записи в таблицах");
            int selectInt = Convert.ToInt32(Console.ReadLine());

            if(selectInt == 1)
            {
                InsertCommand();
            }
            else if (selectInt == 2)
            {
                SelectCommand();
            }
            else if(selectInt == 3)
            {

            }
            

        }

        public static void InsertCommand()
        {
            string CONNECT_STRING = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=5",
               Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database.mdb"));

            OleDbConnection connection = new OleDbConnection(CONNECT_STRING);

            Console.WriteLine("Выберете таблицу для заполнения: 1 - Таблица Производителей, 2 - Таблица моделей, 3 - Таблица причин");
            int selectedInt = Convert.ToInt32(Console.ReadLine());

            if (File.Exists("Database.mdb"))
            {
                try
                {
                    connection.Open();
                    Console.WriteLine("Connection opened");

                    if (selectedInt == 1)
                    {
                        Console.WriteLine("Введите наименование производителя");
                        string strName = Console.ReadLine();

                        string query = "INSERT INTO TablesManufacturer(strManufactName) VALUES ('" + strName + "')";
                       
                        OleDbCommand command = new OleDbCommand(query, connection);

                        command.ExecuteNonQuery();
                    }
                    else if (selectedInt == 2)
                    {
                        Console.WriteLine("Введите наименование модели");
                        string strName = Console.ReadLine();

                        Console.WriteLine("Введите номер семейства моделей");
                        int SMCSFamilyID = Convert.ToInt32(Console.ReadLine());

                        Console.WriteLine("Введите путь к картинке");
                        string strImage = Console.ReadLine();

                        Console.WriteLine("Введите наименование производителя");
                        string strManufactName = Console.ReadLine();

                        string query1 = string.Empty;
                        if (strManufactName == "Apple")
                        {
                            int x = 1;

                            query1 = "INSERT INTO TablesModel(strName, intSMCSFamilyID, strImage, intManufacturerID) VALUES ('" + strName + "', '" + SMCSFamilyID + "', '" + strImage + "', '" + x + "')";
                        }
                        else if(strManufactName == "Sumsyng")
                        {
                            int x = 2;
                            query1 = "INSERT INTO TablesModel(strName, intSMCSFamilyID, strImage, intManufacturerID) VALUES ('" + strName + "', '" + SMCSFamilyID + "', '" + strImage + "', '" + x + "')";
                        }                        
                      
                        OleDbCommand command = new OleDbCommand(query1, connection);
                      
                        command.ExecuteNonQuery();
                    }
                    else if(selectedInt == 3)
                    {
                        Console.WriteLine("Введите причину");
                        string strName = Console.ReadLine();

                        Console.WriteLine("Ответьте true или false");
                        bool bitDowntime = Convert.ToBoolean(Console.ReadLine());

                        Console.WriteLine("Ответьте true или false");
                        bool bitUnscheduled = Convert.ToBoolean(Console.ReadLine());

                        Console.WriteLine("Ответьте true или false");
                        bool bitPMStoppages = Convert.ToBoolean(Console.ReadLine());

                        Console.WriteLine("Ответьте true или false");
                        bool bitScheduledRepairsAndOther = Convert.ToBoolean(Console.ReadLine());

                        Console.WriteLine("Введите номер локации");
                        int intLocationId = Convert.ToInt32(Console.ReadLine());

                        
                        string query1 = "INSERT INTO TablesStopReason(strReason, bitDowntime, bitUnscheduled, bitPMStoppages, bitScheduledRepairsAndOther, intLocationId) VALUES ('" + strName + "', '" + bitDowntime + "', '" + bitUnscheduled + "', '" + bitPMStoppages + "', '" + bitScheduledRepairsAndOther + "', '"+intLocationId+"')";
                       
                        OleDbCommand command = new OleDbCommand(query1, connection);
                       
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    connection.Close();
                    connection.Dispose();
                    Console.WriteLine("Connection closed");
                }

            }

        }

        public static void SelectCommand()
        {
            string CONNECT_STRING = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=5",
               Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database.mdb"));

            OleDbConnection connection = new OleDbConnection(CONNECT_STRING);

            Console.WriteLine("Выберете таблицу для отражения: 1 - Таблица Производителей, 2 - Таблица моделей, 3 - Таблица причин");
            int selectedInt = Convert.ToInt32(Console.ReadLine());

            if (File.Exists("Database.mdb"))
            {
                try
                {
                    connection.Open();
                    Console.WriteLine("Connection opened");

                    if (selectedInt == 1)
                    {     
                        string query = "SELECT * FROM TablesManufacturer";
                       
                        OleDbCommand command = new OleDbCommand(query, connection);

                        OleDbDataReader dr = command.ExecuteReader();
                        while (dr.Read())
                        {
                            string str = string.Format("Manufacturer Name: {0}", dr[1]);

                            Console.WriteLine(str);
                        }

                    }
                    else if (selectedInt == 2)
                    {
                        string query = "SELECT * FROM TablesModel";
                        OleDbCommand command = new OleDbCommand(query, connection);
                        
                        OleDbDataReader dr = command.ExecuteReader();

                        while(dr.Read())
                        {
                            string str = string.Format("Id: {0}, Model Name: {1}, SMCSFamily ID: {2}, Path to Image: {3}, Manufacturer ID: {4}", dr[0], dr[1], dr[2], dr[3], dr[4]);

                            Console.WriteLine(str);
                        }
                    }
                    else if (selectedInt == 3)
                    {
                        string query = "SELECT * FROM TablesStopReason";

                        
                        OleDbCommand command = new OleDbCommand(query, connection);
                        
                        OleDbDataReader dr =  command.ExecuteReader();

                        while(dr.Read())
                        {
                            string str = string.Format("Id: {0}, Reason Name: {1}, bitDowntime: {2}, bitUnscheduled: {3}, bitPMStoppages: {4}, bitScheduledRepairsAndOther: {5}, LocationId: {6}", dr[0], dr[1], dr[2], dr[3], dr[4], dr[5], dr[6]);

                            Console.WriteLine(str);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    connection.Close();
                    connection.Dispose();
                    Console.WriteLine("Connection closed");
                }

            }

        }

        public static void DeleteCommand()
        {
            string CONNECT_STRING = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=5",
               Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database.mdb"));

            OleDbConnection connection = new OleDbConnection(CONNECT_STRING);

            Console.WriteLine("Выберете таблицу для удаления записи: 1 - Таблица Производителей, 2 - Таблица моделей, 3 - Таблица причин");
            int selectedInt = Convert.ToInt32(Console.ReadLine());

            if(selectedInt == 1)
            {
                Console.WriteLine("Введите ID записи");
                int selectedId = Convert.ToInt32(Console.ReadLine());

                string query = "DELETE FROM TablesManufacturer WHERE ID = '" + selectedId + "' ";

                OleDbCommand command = new OleDbCommand(query, connection);
                
                command.ExecuteNonQuery();
            }
            else if (selectedInt == 2)
            {
                Console.WriteLine("Введите ID записи");
                int selectedId = Convert.ToInt32(Console.ReadLine());

                string query = "DELETE FROM TablesModel WHERE ID = '" + selectedId + "' ";

                OleDbCommand command = new OleDbCommand(query, connection);

                command.ExecuteNonQuery();
            }
            else if(selectedInt == 3)
            {
                Console.WriteLine("Введите ID записи");
                int selectedId = Convert.ToInt32(Console.ReadLine());

                string query = "DELETE FROM TablesStopReason WHERE ID = '" + selectedId + "' ";

                OleDbCommand command = new OleDbCommand(query, connection);

                command.ExecuteNonQuery();
            }
        }

    }
        
}


