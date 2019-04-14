using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;

namespace ExcelToSqlBulkCopy
{
    class RelayCommand : ICommand
    {
        private readonly Action<object> _action;

        public event EventHandler CanExecuteChanged = delegate {};

        public RelayCommand(Action<object> action)
        {
            _action = action;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            _action(parameter);
        }
    }
}
