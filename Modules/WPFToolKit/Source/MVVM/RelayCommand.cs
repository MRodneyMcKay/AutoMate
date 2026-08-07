/*This file is part of AutoMate.  

    AutoMate is free software: you can redistribute it and/or modify  
    it under the terms of the GNU General Public License as published by  
    the Free Software Foundation, either version 3 of the License, or  
    (at your option) any later version.  

    This program is distributed in the hope that it will be useful,  
    but WITHOUT ANY WARRANTY; without even the implied warranty of  
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the  
    GNU General Public License for more details.  

    You should have received a copy of the GNU General Public License  
    along with this program. If not, see <https://www.gnu.org/licenses/>.  
*/

using System;
using System.Diagnostics;
using System.Windows.Input;

namespace WPFToolkit.MVVM
{
    public class RelayCommand : ICommand
    {
        private readonly Action<object> execute;
        private readonly Predicate<object> canExecute;


        public RelayCommand(
            Action<object> execute,
            Predicate<object> canExecute = null)
        {
            this.execute = execute;
            this.canExecute = canExecute;
        }


        public bool CanExecute(
            object parameter)
        {
            return canExecute == null ||
                   canExecute(parameter);
        }


        public void Execute(
            object parameter)
        {
            try
            {
                execute(parameter);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                throw;
            }
        }


        public event EventHandler CanExecuteChanged
        {
            add
            {
                CommandManager.RequerySuggested += value;
            }

            remove
            {
                CommandManager.RequerySuggested -= value;
            }
        }
    }
}