using CT;
using e3;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace ConnectToE3
{
    public static class AppConnect
    {
        private static e3Application e3App; // Приложение       
        // Подключение к экземпляру E3
        public static e3Application ToE3()
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
                    e3App = new e3Application();
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
      
        /// <summary>
        /// Перегрузка при подключении к конкретному проекту. prjPath - полный путь к файлу
        /// </summary>
        /// <param name="prjPath"> Путь к файлу</param>
        /// <param name="quitThenDone"> Flag True если нужно будет закрыть приложение </param>
        /// <returns></returns>
        public static e3Application ToE3(string prjPath, out bool quitThenDone)
        {       
            Dispatcher disp = new Dispatcher();
            quitThenDone = false;

            if (disp != null)
            {
                Process[] processList = Process.GetProcessesByName("E3.series"); // получаем процессы E3.series

                foreach (Process process in processList)
                {
                    e3Application App = (e3Application)disp.GetE3ByProcessId(process.Id);
                    if (App == null) continue;   // на случай открытого окна БД, повисших процессов и т.п.

                    e3Job Prj = (e3Job)App.CreateJobObject();
                    string project = Prj.GetPath() + Prj.GetName() + Prj.GetType();
                   // MessageBox.Show(project + "\n" + PrjPath, "Ошибка", MessageBoxButtons.OK);
                    if (string.Equals(project , prjPath, StringComparison.CurrentCultureIgnoreCase))
                    {
                        e3App = App;
                        break;
                    };                    

                };
                // Если не запущенного проекта - запускаем новый процесс
                e3App = (e3Application)disp.OpenE3Application(prjPath);
                quitThenDone = true;
            };

            if (e3App == null)
                MessageBox.Show("Нет e3App.", "Ошибка", MessageBoxButtons.OK);
            return e3App;
        }
        // Получаем список открытых проектов
        public static Dictionary<string, e3Application> GetE3ProcessDictionary()
        {
            Dispatcher disp = new Dispatcher();
            Dictionary<string, e3Application> E3ProcessDictionary = new Dictionary<string, e3Application>();

            if (disp != null)
            {
                Process[] processList = Process.GetProcessesByName("E3.series"); // получаем процессы E3.series

                foreach (Process process in processList)
                {
                    e3Application App = (e3Application)disp.GetE3ByProcessId(process.Id);
                    if (App == null) continue;   // на случай открытого окна БД, повисших процессов и т.п.

                    e3Job Prj = (e3Job)App.CreateJobObject();
                    string project = Prj.GetPath() + Prj.GetName() + Prj.GetType();

                    if (Prj.GetName() == "") continue;   // на случай окна без проекта

                    E3ProcessDictionary.Add(project, App);
                        
                };
            };
          
            return E3ProcessDictionary;
        }
    }
}
