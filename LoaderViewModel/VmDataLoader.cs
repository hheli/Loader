using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using LoaderModel;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows;

namespace LoaderViewModel
{
    class UploadKpfBtmClickCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {

            var bcc = (VmDataLoader)parameter;
            bcc.CallSLXLoader();
        }
    }
    public class VmDataLoader: INotifyPropertyChanged
    {
        private bool _progBarVisibility;
        public event PropertyChangedEventHandler PropertyChanged;
        public string filePath { get; set; }
        public bool isTestMode { get; set; }
        public bool progBarVisibility
        {
            get { return _progBarVisibility; }
            set
            {
                _progBarVisibility = value;
                OnPropertyChanged("progBarVisibility");
            }
        }
        #region RadioButtonVM
        int _Value = 1;
        public int Value
        {
            get { return _Value; }
            set
            {
                //SetProperty(ref _Value, value);
                _Value = value;
                OnPropertyChanged("RbIsOcr");
                OnPropertyChanged("RbIsSzrc");
                OnPropertyChanged("RbIsMbank");
                OnPropertyChanged("RbIsCrm");
                OnPropertyChanged("RbIsOfs");
                OnPropertyChanged("RbIsPS");
                OnPropertyChanged("RbIsGroup");

            }
        }
        public bool RbIsOcr
        {
            get { return Value.Equals(1); }
            set { Value = 1; }
        }
        public bool RbIsSzrc
        {
            get { return Value.Equals(2); }
            set { Value = 2; }
        }

        public bool RbIsMbank
        {
            get { return Value.Equals(3); }
            set { Value = 3; }
        }
        public bool RbIsCrm
        {
            get { return Value.Equals(4); }

            set { Value = 4; }
        }
        public bool RbIsOfs
        {
            get { return Value.Equals(5); }

            set { Value = 5; }
        }
        public bool RbIsPs
        {
            get { return Value.Equals(6); }

            set { Value = 6; }
        }
        public bool RbIsGroup
        {
            get { return Value.Equals(7); }

            set { Value = 7; }
        }

        #endregion RadioButton

        public VmDataLoader()
        {
            filePath = @"C:\";
            isTestMode = false;
            progBarVisibility = false;
        }
        public void CallSLXLoader()
        {
            switch (Value)
            {
                case 1:
                    {
                       
                        var tasks = new Task[2];
                        tasks[0] = Task.Factory.StartNew(() =>
                          {
                              progBarVisibility = true;
                              XlDataLoaderCreator xdlc = new KpfDataLoaderCreator();
                              XlDataLoader xdl = xdlc.CreateXlDataLoader(filePath, isTestMode, "РВПС");
                              xdl.UploadToDb();
                          });
                        tasks[1] = Task.Factory.StartNew(() =>
                        {
                            
                            XlDataLoaderCreator xdlc = new KpfDataLoaderCreator();
                            XlDataLoader xdl = xdlc.CreateXlDataLoader(filePath, isTestMode, "РВП");
                            xdl.UploadToDb();
                        });

                        Task.Factory.ContinueWhenAll(tasks, (Task) =>
                        {
                            if (tasks.All(t => t.Status == TaskStatus.RanToCompletion))
                            {
                                progBarVisibility = false;
                                MessageBox.Show("Task ended.");
                            }
                            else
                            {
                                MessageBox.Show("Что-то не так!!! \n РВПС "+tasks[0].ToString()+", РВП " +tasks[0].ToString());

                            }

                        });
                        break;
                    }
                case 2:
                    {
                        Task t = Task.Factory.StartNew(() => 
                        {
                            progBarVisibility = true;
                            XlDataLoaderCreator xdlc = new SzrcDataLoaderCreator();
                            XlDataLoader xdl = xdlc.CreateXlDataLoader(filePath, isTestMode,"Отчет");
                            xdl.UploadToDb();
                        })
                        .ContinueWith(Task => {  progBarVisibility = false; MessageBox.Show("Task ended."); });
                        break;
                    }
                case 3:
                    {
                        Task t = Task.Factory.StartNew(() =>
                        {
                            progBarVisibility = true;
                            XlDataLoaderCreator xdlc = new MBankDataLoaderCreator();
                            XlDataLoader xdl = xdlc.CreateXlDataLoader(filePath, isTestMode, "По_счетам");
                            xdl.UploadToDb();
                        })
                        .ContinueWith(Task => { progBarVisibility = false; MessageBox.Show("Task ended."); });
                        break;
                    }
                case 4:
                    {
                        Task t = Task.Factory.StartNew(() =>
                        {
                            progBarVisibility = true;
                            CrmDataLoader crmdl = new CrmDataLoader(filePath, isTestMode, "Все клиенты Банка");

                            crmdl.UploadToDb();
                        })
                         .ContinueWith(Task => { progBarVisibility = false; MessageBox.Show("Task ended."); });
                        break;
                    }
                case 5:
                    {
                        Task t = Task.Factory.StartNew(() =>
                        {
                            progBarVisibility = true;
                            OfsDataLoader ofsdl = new OfsDataLoader(filePath, isTestMode, "Лист1");

                            ofsdl.UploadToDb();
                        })
                          .ContinueWith(Task => { progBarVisibility = false; MessageBox.Show("Task ended."); });
                        break;
                    }
                case 6:
                    {
                        Task t = Task.Factory.StartNew(() =>
                        {
                            progBarVisibility = true;
                            PsDataLoader psdl = new PsDataLoader(filePath, isTestMode);
                            psdl.UploadToDb();

                        })
                          .ContinueWith(Task => { progBarVisibility = false; MessageBox.Show("Task ended!"); });
                        break;
                    }
                case 7:
                    {
                        Task t = Task.Factory.StartNew(() =>
                        {
                            progBarVisibility = true;
                            GroupsLoader gl = new GroupsLoader(filePath, isTestMode);

                            gl.UploadToDb();
                        })
                         .ContinueWith(Task => { progBarVisibility = false; MessageBox.Show("Task ended."); });
                        break;
                    }
                default: 
                    {
                        MessageBox.Show("Это невозможно!");
                        break;
                    }
            }
        }
        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }
        public ICommand UploadBtnClickCommand
        {
            get
            {
                return new UploadKpfBtmClickCommand();
            }
        }

    }
}
