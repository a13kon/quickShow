using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using EasyModbus;
using Exceltest.BO;

namespace Exceltest
{
    class Program
    {
        private static object result;


        static void Main(string[] args)
        {
            BO.ExcelHelper helper = new BO.ExcelHelper();

            ModbusClient modbusClient = new ModbusClient("COM10");
            modbusClient.Baudrate = 19200;
            modbusClient.UnitIdentifier = 1;
            modbusClient.UDPFlag = true;
                try
                {
                 

                    
                    {
                        if (helper.Open(@"\\rubus\Сектор Электроники и автоматики\= Шмаглиенко\test.xlsx"))
                        {
                            
                            helper.Set(column: "A", row: 1, data: "=[РЕЕСТР_нестандарта.xlsb]N!$S$1");
                            helper.Save();

                            result = helper.Get(column: "A", row: 1);
                            
                            modbusClient.Connect();
                            modbusClient.WriteSingleRegister(0, Convert.ToInt32(result));
                            
                            Console.WriteLine(result);
                        }
                    }
                
                }
                catch (Exception ex) {
                Console.WriteLine(ex.Message);

            }


            helper.Dispose();
            modbusClient.Disconnect();
            //Console.Read();
            }
        }
    }

