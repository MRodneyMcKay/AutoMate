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

using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace WPFToolkit.MVVM
{
    public class ObservableCollectionEx<T> : ObservableCollection<T>
    {
        public void AddRange(
            IEnumerable<T> items)
        {
            foreach (var item in items)
            {
                Add(item);
            }
        }
    }
}