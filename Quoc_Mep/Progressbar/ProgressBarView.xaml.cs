using System;
using System.Collections.Generic;
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


namespace Quoc_MEP
{
    /// <summary>
    /// Interaction logic for ProgressBarView.xaml
    /// </summary>
    public partial class ProgressBarView : Window, IDisposable
    {
        public bool IsClosed { get; private set; }
        public bool IsCancelled { get; private set; }
        private double maximum;
        private string title;
        public event EventHandler CancelRequested;

        public ProgressBarView(string title, double maximum)
        {
            InitializeComponent();
            this.title = title;
            this.maximum = maximum;
            IsCancelled = false;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsCancelled = true;
            CancelRequested?.Invoke(this, EventArgs.Empty);
            CancelButton.IsEnabled = false;
            CancelButton.Content = "Cancelling...";
        }

        public bool Update (double value = 1)
        {
            DoEvent();
            progressBar.Value += value;

            double percent = Math.Round(progressBar.Value / maximum * 100, 0);
            lb1.Content = percent.ToString() + "%";
            Title = title + progressBar.Value.ToString() + "/" + maximum.ToString();

            return IsClosed;

        }

        private void DoEvent()
        {
            System.Windows.Forms.Application.DoEvents();
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
        }

        public void Dispose()
        {
            if (!IsClosed) Close();
        }
    }
}
