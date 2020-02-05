using CT;
using e3;
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace ConnectToE3
{
    public class AppConnect
    {
        private  e3Application e3App = new e3Application(); // Приложение       
        // Подключение к экземпляру E3
        public e3Application ToE3()
        {
            Dispatcher disp = new Dispatcher();
            DispatcherViewer viewer = new DispatcherViewer();
            if (disp != null)
            {
                Object lst = null;      // Объект списка запущенных экземпляров
                Object e3Obj = new e3Application();

                int ProcessCnt = disp.GetE3Applications(ref lst); // Получаем список запущенных экземпляров
                if (ProcessCnt == 1)
                {
                    return e3App; // Приложение одно 
                }
                else if ((ProcessCnt > 1))
                {
                    if (viewer.ShowViewer(ref e3Obj))
                    {
                        e3App = (e3Application)e3Obj;
                    }
                    else
                    {
                        MessageBox.Show("Нет проекта.", "Ошибка", MessageBoxButtons.OK);
                    };
                };
            };

            if (e3App == null)
                MessageBox.Show("Нет e3App.", "Ошибка", MessageBoxButtons.OK);
            return e3App;
        }
        // Перегрузка при подключении к конкретному проекту. PrjPath - полный путь к файлу
        public e3Application ToE3(string PrjPath)
        {
            Dispatcher disp = new Dispatcher();

            if (disp != null)
            {
                Process[] processList = Process.GetProcessesByName("E3.series"); // получаем процессы E3.series

                foreach (Process process in processList)
                {
                    e3Application App = (e3Application)disp.GetE3ByProcessId(process.Id);
                    if (App == null) continue;   // на случай открытого окна БД, повисших процессов и т.п.

                    e3Job Prj = App.CreateJobObject();
                    string project = Prj.GetPath() + Prj.GetName() + Prj.GetType();

                    if (project == PrjPath)
                    {
                        e3App = App;
                        break;
                    };                    

                };
            };

            if (e3App == null)
                MessageBox.Show("Нет e3App.", "Ошибка", MessageBoxButtons.OK);
            return e3App;
        }
    }
}
