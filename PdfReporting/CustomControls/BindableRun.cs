﻿using System;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Data;

namespace WpfPdfReporting.CustomControls
{
    public class BindableRun : Run
    {
        public static readonly DependencyProperty BoundTextProperty = DependencyProperty.Register("BoundText", typeof(string), typeof(BindableRun), new PropertyMetadata(OnBoundTextChanged));

        public BindableRun()
        {
            Helpers.FixupDataContext(this);
        }

        private static void OnBoundTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            ((Run)d).Text = (string)e.NewValue;
        }

        public string BoundText
        {
            get => (string)GetValue(BoundTextProperty);
            set => SetValue(BoundTextProperty, value);
        }
    }
}
