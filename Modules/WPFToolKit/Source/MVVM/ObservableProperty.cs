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
using System.ComponentModel;

namespace WPFToolkit.MVVM
{
    public class ObservableProperty<T> : INotifyPropertyChanged
    {
        private T value;


        public T Value
        {
            get
            {
                return value;
            }

            set
            {
                if (Equals(this.value, value))
                    return;

                this.value = value;

                PropertyChanged?.Invoke(
                    this,
                    new PropertyChangedEventArgs(nameof(Value))
                );
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;


        public ObservableProperty()
        {
            value = default(T);
        }


        public ObservableProperty(
            T initialValue)
        {
            value = initialValue;
        }
    }
}