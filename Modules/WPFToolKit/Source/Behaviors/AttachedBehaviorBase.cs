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

using System.Windows;

namespace WPFToolkit.Behaviors
{
    public abstract class AttachedBehaviorBase<T>
        where T : AttachedBehaviorBase<T>, new()
    {
        public static readonly DependencyProperty IsEnabledProperty =
            DependencyProperty.RegisterAttached(
                "IsEnabled",
                typeof(bool),
                typeof(T),
                new PropertyMetadata(false, OnIsEnabledChanged)
            );


        public static void SetIsEnabled(
            DependencyObject element,
            bool value)
        {
            element.SetValue(
                IsEnabledProperty,
                value);
        }


        public static bool GetIsEnabled(
            DependencyObject element)
        {
            return (bool)element.GetValue(
                IsEnabledProperty);
        }


        private static void OnIsEnabledChanged(
            DependencyObject sender,
            DependencyPropertyChangedEventArgs e)
        {
            if (sender == null)
                return;


            var behavior = new T();


            if ((bool)e.NewValue)
            {
                behavior.Attach(sender);
            }
            else
            {
                behavior.Detach(sender);
            }
        }


        protected abstract void Attach(
            DependencyObject element);


        protected virtual void Detach(
            DependencyObject element)
        {
        }
    }
}