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
using System.Windows;

namespace WPFToolkit.Behaviors
{
    public static class BehaviorBridge
    {
        public static readonly DependencyProperty BehaviorProperty =
            DependencyProperty.RegisterAttached(
                "Behavior",
                typeof(object),
                typeof(BehaviorBridge),
                new PropertyMetadata(null, OnBehaviorChanged)
            );


        public static void SetBehavior(
            DependencyObject element,
            object value)
        {
            element.SetValue(
                BehaviorProperty,
                value);
        }


        public static object GetBehavior(
            DependencyObject element)
        {
            return element.GetValue(
                BehaviorProperty);
        }


        public static event EventHandler<BehaviorRequestedEventArgs> BehaviorRequested;


        private static void OnBehaviorChanged(
            DependencyObject sender,
            DependencyPropertyChangedEventArgs e)
        {
            if (e.NewValue == null)
                return;


            var behaviors = e.NewValue
                .ToString()
                .Split(',');


            foreach (var behavior in behaviors)
            {
                var name = behavior.Trim();

                if (string.IsNullOrWhiteSpace(name))
                    continue;


                BehaviorRequested?.Invoke(
                    sender,
                    new BehaviorRequestedEventArgs(
                        sender,
                        name
                    ));
            }
        }


        public static readonly DependencyProperty ActionProperty =
            DependencyProperty.RegisterAttached(
                "Action",
                typeof(object),
                typeof(BehaviorBridge),
                new PropertyMetadata(null)
            );


        public static void SetAction(
            DependencyObject element,
            object value)
        {
            element.SetValue(
                ActionProperty,
                value);
        }


        public static object GetAction(
            DependencyObject element)
        {
            return element.GetValue(
                ActionProperty);
        }
    }
}