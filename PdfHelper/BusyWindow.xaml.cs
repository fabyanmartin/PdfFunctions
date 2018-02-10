using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PdfFunctions
{
    /// <summary>
    /// Interaction logic for BusyWindow.xaml
    /// </summary>
    public partial class BusyWindow : Window, INotifyPropertyChanged
    {
        public static readonly DependencyProperty PercentCompleteProperty = DependencyProperty.Register(
           "PercentComplete", typeof(int), typeof(BusyWindow), new PropertyMetadata(0, Callback));

        private static void Callback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var view = d as BusyWindow;
            if (view != null)
            {
                if (e.NewValue != null && (int)e.NewValue > 0)
                {
                    view.CompleteText = e.NewValue.ToString() + " % Complete";
                }
            }
        }

        public int PercentComplete
        {
            get { return (int)GetValue(PercentCompleteProperty); }
            set { SetValue(PercentCompleteProperty, value); }
        }

        private string _completeText;
        public string CompleteText
        {
            get
            {
                return _completeText;
            }
            set
            {
                _completeText = value;
                OnPropertyChanged("CompleteText");
            }
        }


        public BusyWindow()
        {
            InitializeComponent();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
